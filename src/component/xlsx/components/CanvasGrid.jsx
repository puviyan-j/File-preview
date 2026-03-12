import React, { useRef, useEffect, useCallback, useState } from 'react';
import { renderRegion, renderRowHeaders, renderColHeaders } from '../engine/renderEngine.js';
import { getVisibleRange, getContentSize } from '../engine/virtualization.js';
import { createScrollEngine } from '../engine/scrollEngine.js';
import { createTileManager } from '../engine/tileManager.js';

const ROW_HEADER_WIDTH = 50;
const COL_HEADER_HEIGHT = 26;

/**
 * CanvasGrid — High-performance canvas-based spreadsheet grid.
 * Handles scroll, virtualization, tile caching, and rendering.
 */
export default function CanvasGrid({ sheet, workbook, onScrollChange }) {
  const containerRef = useRef(null);
  const mainCanvasRef = useRef(null);
  const rowHeaderCanvasRef = useRef(null);
  const colHeaderCanvasRef = useRef(null);

  const scrollEngineRef = useRef(null);
  const tileManagerRef = useRef(null);
  const rafRef = useRef(null);
  const sizeRef = useRef({ width: 0, height: 0 });
  const scrollRef = useRef({ top: 0, left: 0 });
  const sheetRef = useRef(null);
  const workbookRef = useRef(null);

  // Resizing state
  const resizeRef = useRef({
    active: false,
    type: null, // 'col' or 'row'
    index: -1,
    startX: 0,
    startY: 0,
    startSize: 0,
    currentPos: 0,
  });
  const [resizeLine, setResizeLine] = useState(null);

  // Keep refs in sync
  sheetRef.current = sheet;
  workbookRef.current = workbook;

  // Scrollbar state
  const [scrollbarState, setScrollbarState] = useState({
    vThumbTop: 0, vThumbHeight: 30,
    hThumbLeft: 0, hThumbWidth: 30,
    showV: false, showH: false,
  });

  // Initialize engines
  useEffect(() => {
    tileManagerRef.current = createTileManager();
    scrollEngineRef.current = createScrollEngine({
      onScroll: (top, left) => {
        scrollRef.current = { top, left };
        scheduleRender();
        onScrollChange?.(top, left);
      },
    });

    return () => {
      scrollEngineRef.current?.destroy();
      if (rafRef.current) cancelAnimationFrame(rafRef.current);
    };
  }, []);

  // Register wheel handler as non-passive directly on the DOM element
  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;

    const handler = (e) => {
      scrollEngineRef.current?.handleWheel(e);
    };

    el.addEventListener('wheel', handler, { passive: false });
    return () => el.removeEventListener('wheel', handler);
  }, []);

  // Resize observer
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        const { width, height } = entry.contentRect;
        sizeRef.current = { width, height };
        resizeCanvases(width, height);
        updateContentSize();
        scheduleRender();
      }
    });

    observer.observe(container);
    return () => observer.disconnect();
  }, []);

  // When sheet changes, invalidate and re-render
  useEffect(() => {
    if (!sheet) return;
    tileManagerRef.current?.invalidateAll();
    updateContentSize();
    // Force render by clearing any pending raf and scheduling new one
    if (rafRef.current) {
      cancelAnimationFrame(rafRef.current);
      rafRef.current = null;
    }
    scheduleRender();
  }, [sheet, sheet?.rowCount, sheet?.colCount]);

  // React to progressive rendering updates
  useEffect(() => {
    if (sheet && sheet._renderVersion) {
      tileManagerRef.current?.invalidateAll();
      if (rafRef.current) {
        cancelAnimationFrame(rafRef.current);
        rafRef.current = null;
      }
      scheduleRender();
    }
  }, [sheet?._renderVersion]);

  function updateContentSize() {
    const s = sheetRef.current;
    if (!s || !scrollEngineRef.current) return;
    const { width, height } = getContentSize(s);
    const vp = sizeRef.current;
    scrollEngineRef.current.setContentSize(width, height, vp.width, vp.height);
  }

  function resizeCanvases(w, h) {
    const dpr = window.devicePixelRatio || 1;

    const mainCanvas = mainCanvasRef.current;
    if (mainCanvas) {
      mainCanvas.width = Math.round(w * dpr);
      mainCanvas.height = Math.round(h * dpr);
      mainCanvas.style.width = `${w}px`;
      mainCanvas.style.height = `${h}px`;
    }

    const rowCanvas = rowHeaderCanvasRef.current;
    if (rowCanvas) {
      rowCanvas.width = Math.round(ROW_HEADER_WIDTH * dpr);
      rowCanvas.height = Math.round(h * dpr);
      rowCanvas.style.width = `${ROW_HEADER_WIDTH}px`;
      rowCanvas.style.height = `${h}px`;
    }

    const colCanvas = colHeaderCanvasRef.current;
    if (colCanvas) {
      colCanvas.width = Math.round(w * dpr);
      colCanvas.height = Math.round(COL_HEADER_HEIGHT * dpr);
      colCanvas.style.width = `${w}px`;
      colCanvas.style.height = `${COL_HEADER_HEIGHT}px`;
    }
  }

  // ─── Interaction & Resizing ─────────────────────────────────

  const handleColMouseMove = (e) => {
    if (resizeRef.current.active) return;
    const s = sheetRef.current;
    if (!s) return;

    const rect = colHeaderCanvasRef.current.getBoundingClientRect();
    const x = e.clientX - rect.left + scrollRef.current.left;

    // Find column
    const pos = s.getColPositions();
    let edgeCol = -1;
    for (let c = 0; c < s.colCount; c++) {
      if (s.hiddenCols.has(c)) continue;
      const rightEdge = pos[c] + s.getColWidth(c);
      if (Math.abs(x - rightEdge) < 5) {
        edgeCol = c;
        break;
      }
    }

    if (edgeCol !== -1) {
      document.body.style.cursor = 'col-resize';
    } else {
      document.body.style.cursor = 'default';
    }
  };

  const handleRowMouseMove = (e) => {
    if (resizeRef.current.active) return;
    const s = sheetRef.current;
    if (!s) return;

    const rect = rowHeaderCanvasRef.current.getBoundingClientRect();
    const y = e.clientY - rect.top + scrollRef.current.top;

    const pos = s.getRowPositions();
    let edgeRow = -1;
    for (let r = 0; r < s.rowCount; r++) {
      if (s.hiddenRows.has(r)) continue;
      const bottomEdge = pos[r] + s.getRowHeight(r);
      if (Math.abs(y - bottomEdge) < 5) {
        edgeRow = r;
        break;
      }
    }

    if (edgeRow !== -1) {
      document.body.style.cursor = 'row-resize';
    } else {
      document.body.style.cursor = 'default';
    }
  };

  const handleColMouseDown = (e) => {
    if (document.body.style.cursor !== 'col-resize') return;
    const s = sheetRef.current;
    if (!s) return;

    const rect = colHeaderCanvasRef.current.getBoundingClientRect();
    const x = e.clientX - rect.left + scrollRef.current.left;

    const pos = s.getColPositions();
    let edgeCol = -1;
    for (let c = 0; c < s.colCount; c++) {
      if (s.hiddenCols.has(c)) continue;
      const rightEdge = pos[c] + s.getColWidth(c);
      if (Math.abs(x - rightEdge) < 5) {
        edgeCol = c;
        break;
      }
    }

    if (edgeCol !== -1) {
      resizeRef.current = {
        active: true,
        type: 'col',
        index: edgeCol,
        startX: e.clientX,
        startY: e.clientY,
        startSize: s.getColWidth(edgeCol),
        currentPos: e.clientX - rect.left,
      };
      setResizeLine({ type: 'col', pos: e.clientX - rect.left + ROW_HEADER_WIDTH });
      document.addEventListener('mousemove', handleWindowMouseMove);
      document.addEventListener('mouseup', handleWindowMouseUp);
    }
  };

  const handleRowMouseDown = (e) => {
    if (document.body.style.cursor !== 'row-resize') return;
    const s = sheetRef.current;
    if (!s) return;

    const rect = rowHeaderCanvasRef.current.getBoundingClientRect();
    const y = e.clientY - rect.top + scrollRef.current.top;

    const pos = s.getRowPositions();
    let edgeRow = -1;
    for (let r = 0; r < s.rowCount; r++) {
      if (s.hiddenRows.has(r)) continue;
      const bottomEdge = pos[r] + s.getRowHeight(r);
      if (Math.abs(y - bottomEdge) < 5) {
        edgeRow = r;
        break;
      }
    }

    if (edgeRow !== -1) {
      resizeRef.current = {
        active: true,
        type: 'row',
        index: edgeRow,
        startX: e.clientX,
        startY: e.clientY,
        startSize: s.getRowHeight(edgeRow),
        currentPos: e.clientY - rect.top,
      };
      setResizeLine({ type: 'row', pos: e.clientY - rect.top + COL_HEADER_HEIGHT });
      document.addEventListener('mousemove', handleWindowMouseMove);
      document.addEventListener('mouseup', handleWindowMouseUp);
    }
  };

  const handleWindowMouseMove = (e) => {
    const rs = resizeRef.current;
    if (!rs.active) return;

    if (rs.type === 'col') {
      const delta = e.clientX - rs.startX;
      const newSize = Math.max(10, rs.startSize + delta);
      setResizeLine({ type: 'col', pos: rs.currentPos + (newSize - rs.startSize) + ROW_HEADER_WIDTH });
    } else {
      const delta = e.clientY - rs.startY;
      const newSize = Math.max(10, rs.startSize + delta);
      setResizeLine({ type: 'row', pos: rs.currentPos + (newSize - rs.startSize) + COL_HEADER_HEIGHT });
    }
  };

  const handleWindowMouseUp = (e) => {
    const rs = resizeRef.current;
    if (!rs.active) return;

    const s = sheetRef.current;
    if (s) {
      if (rs.type === 'col') {
        const delta = e.clientX - rs.startX;
        const newSize = Math.max(10, rs.startSize + delta);
        s.setColWidth(rs.index, newSize);
      } else {
        const delta = e.clientY - rs.startY;
        const newSize = Math.max(10, rs.startSize + delta);
        s.setRowHeight(rs.index, newSize);
      }
      s.invalidatePositionCache();
      updateContentSize();
      scheduleRender();
    }

    resizeRef.current.active = false;
    setResizeLine(null);
    document.body.style.cursor = 'default';
    document.removeEventListener('mousemove', handleWindowMouseMove);
    document.removeEventListener('mouseup', handleWindowMouseUp);
  };

  useEffect(() => {
    return () => {
      document.removeEventListener('mousemove', handleWindowMouseMove);
      document.removeEventListener('mouseup', handleWindowMouseUp);
      document.body.style.cursor = 'default';
    };
  }, []);

  function scheduleRender() {
    if (rafRef.current) return;
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = null;
      doRender();
    });
  }

  function doRender() {
    const currentSheet = sheetRef.current;
    const currentWorkbook = workbookRef.current;
    if (!currentSheet || !currentWorkbook) return;

    const { width, height } = sizeRef.current;
    if (width === 0 || height === 0) return;

    const { top: scrollTop, left: scrollLeft } = scrollRef.current;
    const dpr = window.devicePixelRatio || 1;

    // Compute visible range
    const range = getVisibleRange({
      scrollTop,
      scrollLeft,
      viewportWidth: width,
      viewportHeight: height,
      sheet: currentSheet,
      overscan: 3,
    });

    // Main grid
    const mainCtx = mainCanvasRef.current?.getContext('2d');
    if (mainCtx) {
      renderRegion(mainCtx, {
        startRow: range.startRow,
        endRow: range.endRow,
        startCol: range.startCol,
        endCol: range.endCol,
        offsetX: scrollLeft,
        offsetY: scrollTop,
        sheet: currentSheet,
        workbook: currentWorkbook,
        width,
        height,
        dpr,
      });
    }

    // Row headers
    const rowCtx = rowHeaderCanvasRef.current?.getContext('2d');
    if (rowCtx) {
      renderRowHeaders(rowCtx, {
        startRow: range.startRow,
        endRow: range.endRow,
        offsetY: scrollTop,
        width: ROW_HEADER_WIDTH,
        height,
        sheet: currentSheet,
        dpr,
      });
    }

    // Column headers
    const colCtx = colHeaderCanvasRef.current?.getContext('2d');
    if (colCtx) {
      renderColHeaders(colCtx, {
        startCol: range.startCol,
        endCol: range.endCol,
        offsetX: scrollLeft,
        width,
        height: COL_HEADER_HEIGHT,
        sheet: currentSheet,
        dpr,
      });
    }

    // Update scrollbar thumbs
    updateScrollbars(scrollTop, scrollLeft, currentSheet);
  }

  function updateScrollbars(scrollTop, scrollLeft, currentSheet) {
    if (!currentSheet) return;
    const { width, height } = sizeRef.current;
    const contentSize = getContentSize(currentSheet);

    if (contentSize.height === 0 || contentSize.width === 0) return;

    const vRatio = height / contentSize.height;
    const hRatio = width / contentSize.width;

    const showV = vRatio < 1;
    const showH = hRatio < 1;

    const vThumbHeight = Math.max(30, height * vRatio);
    const vThumbTop = contentSize.height > 0
      ? (scrollTop / contentSize.height) * (height - vThumbHeight)
      : 0;

    const hThumbWidth = Math.max(30, width * hRatio);
    const hThumbLeft = contentSize.width > 0
      ? (scrollLeft / contentSize.width) * (width - hThumbWidth)
      : 0;

    setScrollbarState({ vThumbTop, vThumbHeight, hThumbLeft, hThumbWidth, showV, showH });
  }

  return (
    <div className="sv-grid-area">
      {/* Corner cell */}
      <div className="sv-corner" />

      {/* Column headers */}
      <div
        className="sv-col-headers"
        // onMouseMove={handleColMouseMove}
        onMouseLeave={() => { if (!resizeRef.current.active) document.body.style.cursor = 'default'; }}
      // onMouseDown={handleColMouseDown}
      >
        <canvas ref={colHeaderCanvasRef} />
      </div>

      {/* Row headers */}
      <div
        className="sv-row-headers"
        // onMouseMove={handleRowMouseMove}
        onMouseLeave={() => { if (!resizeRef.current.active) document.body.style.cursor = 'default'; }}
      // onMouseDown={handleRowMouseDown}
      >
        <canvas ref={rowHeaderCanvasRef} />
      </div>

      {/* Resize Line Guide */}
      {resizeLine && (
        <div style={{
          position: 'absolute',
          background: '#1a73e8',
          zIndex: 10,
          pointerEvents: 'none',
          ...(resizeLine.type === 'col' ? {
            top: 0,
            bottom: 0,
            left: `${resizeLine.pos}px`,
            width: '2px'
          } : {
            left: 0,
            right: 0,
            top: `${resizeLine.pos}px`,
            height: '2px'
          })
        }} />
      )}

      {/* Main grid viewport */}
      <div
        className="sv-grid-viewport"
        ref={containerRef}
      >
        <canvas ref={mainCanvasRef} className="sv-grid-canvas" />
      </div>

      {/* Vertical scrollbar */}
      {scrollbarState.showV && (
        <div className="sv-scrollbar-v sv-scrollbar-active">
          <div
            className="sv-scrollbar-v-thumb"
            style={{
              top: `${scrollbarState.vThumbTop}px`,
              height: `${scrollbarState.vThumbHeight}px`,
            }}
          />
        </div>
      )}

      {/* Horizontal scrollbar */}
      {scrollbarState.showH && (
        <div className="sv-scrollbar-h sv-scrollbar-active">
          <div
            className="sv-scrollbar-h-thumb"
            style={{
              left: `${scrollbarState.hThumbLeft}px`,
              width: `${scrollbarState.hThumbWidth}px`,
            }}
          />
        </div>
      )}
    </div>
  );
}
