import React, { useState, useRef, useCallback, useEffect } from 'react';
import './SheetViewer.css';
import { createWorkbookModel } from '../engine/workbookModel.js';
import CanvasGrid from './CanvasGrid.jsx';
import SheetTabs from './SheetTabs.jsx';
import OverlayLayer from './OverlayLayer.jsx';

// Import the parser worker
import ParserWorker from '../workers/parser.worker.js?worker';

// Supported formats
const SUPPORTED_EXTENSIONS = ['xlsx', 'xls', 'ods', 'csv', 'tsv'];
const ACCEPT_STRING = SUPPORTED_EXTENSIONS.map(e => `.${e}`).join(',');

/**
 * SheetViewer — Main container component.
 * Handles file input, progressive parsing, chunk processing, and rendering orchestration.
 */
export default function SheetViewer() {
  const [workbook, setWorkbook] = useState(null);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [parsing, setParsing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [fileName, setFileName] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [scrollPos, setScrollPos] = useState({ top: 0, left: 0 });
  const [lastUpdate, setLastUpdate] = useState(0);

  const workbookRef = useRef(null);
  const workerRef = useRef(null);
  const chunkQueueRef = useRef([]);
  const processingRef = useRef(false);
  const fileInputRef = useRef(null);
  const renderVersionRef = useRef(0);

  // Auto-load from URL query parameter: ?file=test_data.xlsx
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const autoFile = params.get('file');
    if (autoFile) {
      fetch(`/${autoFile}`)
        .then(res => {
          if (!res.ok) throw new Error(`Failed to fetch ${autoFile}: ${res.status}`);
          return res.arrayBuffer();
        })
        .then(buffer => {
          setFileName(autoFile);
          setParsing(true);
          setProgress(0);
          setActiveSheetIndex(0);
          const wb = createWorkbookModel();
          workbookRef.current = wb;
          setWorkbook(wb);
          startParsing(buffer, autoFile);
        })
        .catch(err => console.error('[AutoLoad] Error:', err));
    }
  }, []);

  // Memory profiling
  useEffect(() => {
    if (!workbook) return;
    const interval = setInterval(() => {
      logMemoryUsage();
    }, 10000);
    return () => clearInterval(interval);
  }, [workbook]);

  function logMemoryUsage() {
    if (performance.memory) {
      const used = (performance.memory.usedJSHeapSize / 1048576).toFixed(1);
      const total = (performance.memory.totalJSHeapSize / 1048576).toFixed(1);
      const limit = (performance.memory.jsHeapSizeLimit / 1048576).toFixed(1);
      console.log(`[Memory] Used: ${used}MB / Total: ${total}MB / Limit: ${limit}MB`);
    }
  }

  // ─── File Input ─────────────────────────────────────────────
  const handleFile = useCallback((file) => {
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!SUPPORTED_EXTENSIONS.includes(ext)) {
      alert(`Unsupported format: .${ext}\nSupported: ${SUPPORTED_EXTENSIONS.join(', ')}`);
      return;
    }

    setFileName(file.name);
    setParsing(true);
    setProgress(0);
    setActiveSheetIndex(0);

    // Create fresh workbook model
    const wb = createWorkbookModel();
    workbookRef.current = wb;
    setWorkbook(wb);

    // Read file as ArrayBuffer
    const reader = new FileReader();
    reader.onload = (e) => {
      startParsing(e.target.result, file.name);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleFileInput = useCallback((e) => {
    handleFile(e.target.files?.[0]);
  }, [handleFile]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    handleFile(e.dataTransfer.files?.[0]);
  }, [handleFile]);

  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setDragOver(false);
  }, []);

  // ─── Parser Worker ─────────────────────────────────────────
  function startParsing(buffer, name) {
    // Terminate previous worker
    workerRef.current?.terminate();

    const worker = new ParserWorker();
    workerRef.current = worker;
    chunkQueueRef.current = [];

    worker.onmessage = (e) => {
      const msg = e.data;

      switch (msg.type) {
        case 'PARSING_START':
          console.log(`[Parser] Started parsing: ${msg.fileName}`);
          break;

        case 'SHEET_META':
          handleSheetMeta(msg);
          break;

        case 'STYLES':
          handleStyles(msg.styles);
          break;

        case 'ROW_CHUNK':
          // Push to queue for idle processing
          chunkQueueRef.current.push(msg);
          scheduleChunkProcessing();
          break;

        case 'PARSING_PROGRESS':
          setProgress(msg.progress);
          break;

        case 'PARSING_COMPLETE':
          console.log('[Parser] Parsing complete');
          workbookRef.current?.setParsingComplete();
          setParsing(false);
          setProgress(1);
          // Process any remaining chunks
          processAllChunks();
          break;

        case 'SHEET_EMPTY':
          workbookRef.current?.getOrCreateSheet(msg.sheet);
          break;
      }
    };

    worker.onerror = (err) => {
      console.error('[Parser] Worker error:', err);
      setParsing(false);
    };

    // Send file to worker
    worker.postMessage({ type: 'PARSE', buffer, fileName: name }, [buffer]);
  }

  function handleSheetMeta(msg) {
    const wb = workbookRef.current;
    if (!wb) return;
    wb.applySheetMetadata(msg.sheet, msg.meta);
    triggerUpdate();
  }

  function handleStyles(styles) {
    const wb = workbookRef.current;
    if (!wb) return;
    wb.applyStyles(styles);
  }

  // ─── Chunk Processing ──────────────────────────────────────
  function scheduleChunkProcessing() {
    if (processingRef.current) return;
    processingRef.current = true;

    if (typeof requestIdleCallback === 'function') {
      requestIdleCallback(processChunkBatch, { timeout: 100 });
    } else {
      setTimeout(processChunkBatch, 0);
    }
  }

  function processChunkBatch(deadline) {
    const queue = chunkQueueRef.current;
    const wb = workbookRef.current;
    if (!wb || queue.length === 0) {
      processingRef.current = false;
      return;
    }

    const hasDeadline = deadline && typeof deadline.timeRemaining === 'function';
    let processed = 0;

    while (queue.length > 0) {
      // check time budget
      if (hasDeadline && deadline.timeRemaining() < 2 && processed > 0) break;

      const chunk = queue.shift();
      wb.addChunk(chunk.sheet, chunk.startRow, chunk.startCol || 0, chunk.rows, chunk.styleRows);
      processed++;
    }

    if (processed > 0) {
      triggerUpdate();
    }

    if (queue.length > 0) {
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(processChunkBatch, { timeout: 100 });
      } else {
        setTimeout(processChunkBatch, 0);
      }
    } else {
      processingRef.current = false;
    }
  }

  function processAllChunks() {
    const queue = chunkQueueRef.current;
    const wb = workbookRef.current;
    if (!wb) return;

    while (queue.length > 0) {
      const chunk = queue.shift();
      wb.addChunk(chunk.sheet, chunk.startRow, chunk.startCol || 0, chunk.rows, chunk.styleRows);
    }
    processingRef.current = false;
    triggerUpdate();
  }

  function triggerUpdate() {
    renderVersionRef.current++;
    setLastUpdate(renderVersionRef.current);
  }

  // ─── Derived state ─────────────────────────────────────────
  const activeSheet = workbook?.getSheet(activeSheetIndex);
  if (activeSheet) {
    activeSheet._renderVersion = lastUpdate;
  }

  const sheetInfo = workbook
    ? {
        rows: activeSheet?.rowCount || 0,
        cols: activeSheet?.colCount || 0,
        sheets: workbook.sheetCount,
        strings: workbook.stringPool.size,
      }
    : null;

  // ─── Render ─────────────────────────────────────────────────
  if (!workbook) {
    return (
      <div className="sv-root">
        <div className="sv-toolbar">
          <div className="sv-toolbar-title">📊 Spreadsheet Preview</div>
        </div>
        <div
          className={`sv-dropzone ${dragOver ? 'sv-dragover' : ''}`}
          onClick={() => fileInputRef.current?.click()}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept={ACCEPT_STRING}
            onChange={handleFileInput}
            style={{ display: 'none' }}
          />
          <div className="sv-dropzone-icon">
            <svg viewBox="0 0 24 24">
              <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8l-6-6zm-1 2l5 5h-5V4zM6 20V4h5v7h7v9H6z" />
              <path d="M8 13h8v1H8zm0 3h6v1H8z" />
            </svg>
          </div>
          <div className="sv-dropzone-text">
            <h3>Drop a spreadsheet file here</h3>
            <p>or click to browse</p>
          </div>
          <div className="sv-dropzone-formats">
            {SUPPORTED_EXTENSIONS.map(ext => (
              <span key={ext}>.{ext.toUpperCase()}</span>
            ))}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="sv-root">
      {/* Toolbar */}
      <div className="sv-toolbar">
        <div className="sv-toolbar-title">📊 {fileName}</div>
        <div className="sv-toolbar-info">
          {sheetInfo && (
            <>
              <span>{sheetInfo.rows.toLocaleString()} rows</span>
              <span>{sheetInfo.cols} columns</span>
              <span>{sheetInfo.sheets} sheet{sheetInfo.sheets > 1 ? 's' : ''}</span>
            </>
          )}
          {parsing && <span>⏳ Parsing...</span>}
          <span
            style={{ cursor: 'pointer', opacity: 0.85 }}
            onClick={() => {
              setWorkbook(null);
              workbookRef.current = null;
              setFileName('');
              setParsing(false);
              setProgress(0);
              workerRef.current?.terminate();
            }}
            title="Open another file"
          >
            ✕ Close
          </span>
        </div>
      </div>

      {/* Progress bar */}
      {parsing && (
        <div className="sv-progress-bar">
          <div
            className="sv-progress-fill"
            style={{ width: `${Math.round(progress * 100)}%` }}
          />
        </div>
      )}

      {/* Canvas Grid */}
      <CanvasGrid
        sheet={activeSheet}
        workbook={workbook}
        onScrollChange={(top, left) => setScrollPos({ top, left })}
      />

      {/* Overlay (images, charts) */}
      {activeSheet && (
        <OverlayLayer
          images={activeSheet.images}
          charts={activeSheet.charts}
          scrollTop={scrollPos.top}
          scrollLeft={scrollPos.left}
          sheet={activeSheet}
        />
      )}

      {/* Sheet Tabs */}
      <SheetTabs
        sheets={workbook.sheets}
        activeIndex={activeSheetIndex}
        onSheetChange={setActiveSheetIndex}
      />

      {/* Status Bar */}
      <div className="sv-status">
        <span className="sv-status-item">
          Sheet: {activeSheet?.name || '—'}
        </span>
        {sheetInfo && (
          <>
            <span className="sv-status-item">
              {sheetInfo.rows.toLocaleString()} × {sheetInfo.cols}
            </span>
            <span className="sv-status-item">
              Strings: {sheetInfo.strings.toLocaleString()}
            </span>
          </>
        )}
        {parsing && (
          <span className="sv-status-item">
            Loading... {Math.round(progress * 100)}%
          </span>
        )}
      </div>
    </div>
  );
}
