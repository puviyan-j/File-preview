/**
 * Canvas Render Engine
 *
 * Draws spreadsheet cells onto a canvas context.
 * Multi-pass: backgrounds → cell backgrounds → borders → text.
 */

import { DEFAULT_ROW_HEIGHT, DEFAULT_COL_WIDTH, SCALE_FACTOR } from './workbookModel.js';

// ─── Colors ───────────────────────────────────────────────────
const GRID_COLOR = '#d0d7de';
const HEADER_BG = '#f6f8fa';
const HEADER_BORDER = '#d0d7de';
const HEADER_TEXT = '#57606a';
const DEFAULT_TEXT_COLOR = '#1f2328';

// Border style weight map (Excel border style → pixel width)
const BORDER_WIDTHS = {
  thin: 1,
  medium: 2,
  thick: 3,
  dashed: 1,
  dotted: 1,
  double: 2,
  hair: 0.5,
  mediumDashed: 2,
  dashDot: 1,
  mediumDashDot: 2,
  dashDotDot: 1,
  mediumDashDotDot: 2,
  slantDashDot: 1,
};

// ─── Helpers ──────────────────────────────────────────────────

/**
 * Compute the pixel rect for a cell (or merged cell range).
 */
function getCellRect(r, c, sheet, colPositions, rowPositions, offsetX, offsetY) {
  const merge = sheet.getMerge(r, c);
  if (merge && merge.startRow === r && merge.startCol === c) {
    const x = (colPositions[merge.startCol] || 0) - offsetX;
    const y = (rowPositions[merge.startRow] || 0) - offsetY;
    // Safe right edge: use position of endCol+1 if available, else accumulate
    const endColNext = merge.endCol + 1;
    const x2 = endColNext < colPositions.length
      ? colPositions[endColNext] - offsetX
      : (colPositions[merge.endCol] || 0) - offsetX + sheet.getColWidth(merge.endCol);
    const endRowNext = merge.endRow + 1;
    const y2 = endRowNext < rowPositions.length
      ? rowPositions[endRowNext] - offsetY
      : (rowPositions[merge.endRow] || 0) - offsetY + sheet.getRowHeight(merge.endRow);
    return { x, y, w: x2 - x, h: y2 - y, isMerge: true };
  }
  const x = (colPositions[c] || 0) - offsetX;
  const y = (rowPositions[r] || 0) - offsetY;
  return { x, y, w: sheet.getColWidth(c), h: sheet.getRowHeight(r), isMerge: false };
}

/**
 * Resolve a raw cell value stored in the column store.
 * Numbers that are valid string-pool IDs are resolved as strings.
 */
function resolveText(raw, stringPool) {
  if (raw === undefined || raw === null) return null;
  if (typeof raw === 'number') {
    if (Number.isInteger(raw) && raw >= 0 && raw < stringPool.size) {
      const s = stringPool.get(raw);
      if (s !== '') return s;
    }
    // It's actually a numeric value
    return String(raw);
  }
  return String(raw);
}

// ─── Main Render Region ───────────────────────────────────────

/**
 * Render a rectangular region of the sheet onto a canvas context.
 */
export function renderRegion(ctx, params) {
  const {
    startRow, endRow, startCol, endCol,
    offsetX, offsetY,
    sheet, workbook,
    width, height,
    dpr = 1,
  } = params;

  if (!sheet || !workbook) return;

  const rowPositions = sheet.getRowPositions();
  const colPositions = sheet.getColPositions();
  const stringPool = workbook.stringPool;
  const stylePool = workbook.stylePool;

  ctx.save();
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, width, height);

  // ── Pass 1: White base + Grid lines ─────────────────────────
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);

  ctx.strokeStyle = GRID_COLOR;
  ctx.lineWidth = 0.5;
  ctx.beginPath();

  for (let r = startRow; r <= endRow && r < rowPositions.length; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    const y = Math.round(rowPositions[r] - offsetY) + 0.5;
    if (y < -1 || y > height + 1) continue;
    ctx.moveTo(0, y);
    ctx.lineTo(width, y);
  }
  for (let c = startCol; c <= endCol && c < colPositions.length; c++) {
    if (sheet.hiddenCols.has(c)) continue;
    const x = Math.round(colPositions[c] - offsetX) + 0.5;
    if (x < -1 || x > width + 1) continue;
    ctx.moveTo(x, 0);
    ctx.lineTo(x, height);
  }
  ctx.stroke();

  // ── Pass 2: Cell backgrounds (including merged cells) ────────
  // First, handle all merge origins visible in viewport (scan all sheet merges)
  const renderedMerges = new Set();
  for (const merge of sheet.merges) {
    // Only render if merge overlaps with the visible viewport area
    const mx1 = (colPositions[merge.startCol] || 0) - offsetX;
    const mx2 = merge.endCol + 1 < colPositions.length
      ? colPositions[merge.endCol + 1] - offsetX
      : (colPositions[merge.endCol] || 0) - offsetX + sheet.getColWidth(merge.endCol);
    const my1 = (rowPositions[merge.startRow] || 0) - offsetY;
    const my2 = merge.endRow + 1 < rowPositions.length
      ? rowPositions[merge.endRow + 1] - offsetY
      : (rowPositions[merge.endRow] || 0) - offsetY + sheet.getRowHeight(merge.endRow);

    if (mx2 < 0 || mx1 > width || my2 < 0 || my1 > height) continue;

    const key = `${merge.startRow},${merge.startCol}`;
    if (renderedMerges.has(key)) continue;
    renderedMerges.add(key);

    const styleId = sheet.store.getCellStyle(merge.startCol, merge.startRow);
    const style = stylePool.get(styleId);

    // Always paint to cover grid lines
    ctx.fillStyle = (style && style.backgroundColor) ? style.backgroundColor : '#ffffff';
    ctx.fillRect(mx1, my1, mx2 - mx1, my2 - my1);
  }

  // Then regular (non-merged) cell backgrounds
  for (let r = startRow; r < endRow; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    for (let c = startCol; c < endCol; c++) {
      if (sheet.hiddenCols.has(c)) continue;

      const merge = sheet.getMerge(r, c);
      // Skip cells that are part of a merge (already handled above)
      if (merge) continue;

      const styleId = sheet.store.getCellStyle(c, r);
      if (!styleId) continue;
      const style = stylePool.get(styleId);
      if (!style || !style.backgroundColor) continue;

      const x = (colPositions[c] || 0) - offsetX;
      const y = (rowPositions[r] || 0) - offsetY;
      const w = sheet.getColWidth(c);
      const h = sheet.getRowHeight(r);

      ctx.fillStyle = style.backgroundColor;
      ctx.fillRect(x, y, w, h);
    }
  }

  // ── Pass 3: Cell Text ─────────────────────────────────────────
  const renderedMergeText = new Set();

  for (let r = startRow; r < endRow; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    for (let c = startCol; c < endCol; c++) {
      if (sheet.hiddenCols.has(c)) continue;

      const merge = sheet.getMerge(r, c);
      if (merge) {
        // Only render text at the merge origin
        if (merge.startRow !== r || merge.startCol !== c) continue;
        const mKey = `${merge.startRow},${merge.startCol}`;
        if (renderedMergeText.has(mKey)) continue;
        renderedMergeText.add(mKey);
      }

      const raw = sheet.store.getCell(c, r);
      // raw === 0 is a valid numeric value — only skip truly absent cells
      if (raw === undefined || raw === null) continue;

      const text = resolveText(raw, stringPool);
      if (!text || text === 'null' || text === 'undefined') continue;

      const styleId = sheet.store.getCellStyle(c, r);
      const style = stylePool.get(styleId);
      const defaultStyle = stylePool.defaultStyle;

      // Font
      const baseFontSize = style.fontSize || defaultStyle.fontSize || 11;
      const fontSize = baseFontSize * SCALE_FACTOR;
      const fontWeight = style.fontWeight || defaultStyle.fontWeight;
      const fontStyle = style.fontStyle || defaultStyle.fontStyle;
      const fontFamily = style.fontFamily || defaultStyle.fontFamily;
      ctx.font = `${fontStyle} ${fontWeight} ${fontSize}px "${fontFamily}", Arial, sans-serif`;
      ctx.fillStyle = style.color || defaultStyle.color || DEFAULT_TEXT_COLOR;

      // Cell rect
      let x, y, cellW, cellH;
      if (merge) {
        const endColNext = merge.endCol + 1;
        const x2 = endColNext < colPositions.length
          ? colPositions[endColNext] - offsetX
          : (colPositions[merge.endCol] || 0) - offsetX + sheet.getColWidth(merge.endCol);
        const endRowNext = merge.endRow + 1;
        const y2 = endRowNext < rowPositions.length
          ? rowPositions[endRowNext] - offsetY
          : (rowPositions[merge.endRow] || 0) - offsetY + sheet.getRowHeight(merge.endRow);
        x = (colPositions[merge.startCol] || 0) - offsetX;
        y = (rowPositions[merge.startRow] || 0) - offsetY;
        cellW = x2 - x;
        cellH = y2 - y;
      } else {
        x = (colPositions[c] || 0) - offsetX;
        y = (rowPositions[r] || 0) - offsetY;
        cellW = sheet.getColWidth(c);
        cellH = sheet.getRowHeight(r);
      }

      const padding = 4;

      // Horizontal alignment
      const textAlign = style.textAlign || 'left';
      let textX;
      if (textAlign === 'center') {
        ctx.textAlign = 'center';
        textX = x + cellW / 2;
      } else if (textAlign === 'right') {
        ctx.textAlign = 'right';
        textX = x + cellW - padding;
      } else {
        ctx.textAlign = 'left';
        textX = x + padding;
      }

      // Vertical alignment
      const vertAlign = style.verticalAlign || defaultStyle.verticalAlign;
      let textY;
      if (vertAlign === 'top') {
        ctx.textBaseline = 'top';
        textY = y + padding;
      } else if (vertAlign === 'center' || vertAlign === 'middle') {
        ctx.textBaseline = 'middle';
        textY = y + cellH / 2;
      } else {
        // Excel default: bottom
        ctx.textBaseline = 'alphabetic';
        textY = y + cellH - padding;
      }

      // Clip text to cell bounds
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, cellW, cellH);
      ctx.clip();
      ctx.fillText(text, textX, textY);
      ctx.restore();
    }
  }

  // ── Pass 4: Borders ───────────────────────────────────────────
  for (let r = startRow; r < endRow; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    for (let c = startCol; c < endCol; c++) {
      if (sheet.hiddenCols.has(c)) continue;

      const merge = sheet.getMerge(r, c);
      // For merged cells, only draw borders at origin
      if (merge && (merge.startRow !== r || merge.startCol !== c)) continue;

      const styleId = sheet.store.getCellStyle(c, r);
      if (!styleId) continue;
      const style = stylePool.get(styleId);
      if (!style) continue;

      let x, y, cellW, cellH;
      if (merge) {
        const endColNext = merge.endCol + 1;
        const x2 = endColNext < colPositions.length
          ? colPositions[endColNext] - offsetX
          : (colPositions[merge.endCol] || 0) - offsetX + sheet.getColWidth(merge.endCol);
        const endRowNext = merge.endRow + 1;
        const y2 = endRowNext < rowPositions.length
          ? rowPositions[endRowNext] - offsetY
          : (rowPositions[merge.endRow] || 0) - offsetY + sheet.getRowHeight(merge.endRow);
        x = (colPositions[merge.startCol] || 0) - offsetX;
        y = (rowPositions[merge.startRow] || 0) - offsetY;
        cellW = x2 - x;
        cellH = y2 - y;
      } else {
        x = (colPositions[c] || 0) - offsetX;
        y = (rowPositions[r] || 0) - offsetY;
        cellW = sheet.getColWidth(c);
        cellH = sheet.getRowHeight(r);
      }

      drawCellBorders(ctx, style, x, y, cellW, cellH);
    }
  }

  ctx.restore();
}

function drawCellBorders(ctx, style, x, y, w, h) {
  const sides = [
    { key: 'borderTop', x1: x, y1: y, x2: x + w, y2: y },
    { key: 'borderBottom', x1: x, y1: y + h, x2: x + w, y2: y + h },
    { key: 'borderLeft', x1: x, y1: y, x2: x, y2: y + h },
    { key: 'borderRight', x1: x + w, y1: y, x2: x + w, y2: y + h },
  ];

  for (const { key, x1, y1, x2, y2 } of sides) {
    const border = style[key];
    if (!border) continue;

    const lw = BORDER_WIDTHS[border.style] || 1;
    ctx.strokeStyle = border.color || '#000000';
    ctx.lineWidth = lw;

    if (border.style === 'dashed' || border.style === 'mediumDashed') {
      ctx.setLineDash([4, 2]);
    } else if (border.style === 'dotted') {
      ctx.setLineDash([1, 2]);
    } else if (border.style === 'dashDot' || border.style === 'mediumDashDot') {
      ctx.setLineDash([4, 2, 1, 2]);
    } else {
      ctx.setLineDash([]);
    }

    ctx.beginPath();
    ctx.moveTo(Math.round(x1) + 0.5, Math.round(y1) + 0.5);
    ctx.lineTo(Math.round(x2) + 0.5, Math.round(y2) + 0.5);
    ctx.stroke();
  }

  // Reset
  ctx.setLineDash([]);
  ctx.lineWidth = 1;
}

// ─── Row / Column Headers ─────────────────────────────────────

export function renderRowHeaders(ctx, { startRow, endRow, offsetY, width, height, sheet, dpr = 1 }) {
  const rowPositions = sheet.getRowPositions();

  ctx.save();
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, width, height);

  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, width, height);

  // Right border line
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(width - 0.5, 0);
  ctx.lineTo(width - 0.5, height);
  ctx.stroke();

  ctx.fillStyle = HEADER_TEXT;
  ctx.font = '11px "Inter", "Segoe UI", Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  for (let r = startRow; r < endRow && r < rowPositions.length; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    const y = rowPositions[r] - offsetY;
    const h = sheet.getRowHeight(r);
    if (y + h < 0 || y > height) continue;

    ctx.fillText(String(r + 1), width / 2, y + h / 2);

    ctx.beginPath();
    ctx.moveTo(0, Math.round(y + h) + 0.5);
    ctx.lineTo(width, Math.round(y + h) + 0.5);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.stroke();
  }

  ctx.restore();
}

export function renderColHeaders(ctx, { startCol, endCol, offsetX, width, height, sheet, dpr = 1 }) {
  const colPositions = sheet.getColPositions();

  ctx.save();
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, width, height);

  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, width, height);

  // Bottom border
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(0, height - 0.5);
  ctx.lineTo(width, height - 0.5);
  ctx.stroke();

  ctx.fillStyle = HEADER_TEXT;
  ctx.font = '11px "Inter", "Segoe UI", Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  for (let c = startCol; c < endCol && c < colPositions.length; c++) {
    if (sheet.hiddenCols.has(c)) continue;
    const x = colPositions[c] - offsetX;
    const w = sheet.getColWidth(c);
    if (x + w < 0 || x > width) continue;

    ctx.fillText(colIndexToLetter(c), x + w / 2, height / 2);

    ctx.beginPath();
    ctx.moveTo(Math.round(x + w) + 0.5, 0);
    ctx.lineTo(Math.round(x + w) + 0.5, height);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.stroke();
  }

  ctx.restore();
}

function colIndexToLetter(index) {
  let result = '';
  let n = index;
  while (n >= 0) {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  }
  return result;
}
