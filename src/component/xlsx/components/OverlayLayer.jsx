import React from 'react';

/**
 * OverlayLayer — Positions images and charts over the canvas grid using DOM elements.
 */
export default function OverlayLayer({ images, charts, scrollTop, scrollLeft, sheet }) {
  if ((!images || images.length === 0) && (!charts || charts.length === 0)) {
    return null;
  }

  const rowPositions = sheet?.getRowPositions();
  const colPositions = sheet?.getColPositions();

  return (
    <div className="sv-overlay">
      {/* Images */}
      {images && images.map((img, idx) => {
        if (!rowPositions || !colPositions) return null;
        const top = (rowPositions[img.row] || 0) - scrollTop;
        const left = (colPositions[img.col] || 0) - scrollLeft;

        return (
          <img
            key={`img-${idx}`}
            className="sv-overlay-image"
            src={img.src}
            alt={`Image at ${img.row},${img.col}`}
            style={{
              top: `${top}px`,
              left: `${left}px`,
              width: img.width ? `${img.width}px` : 'auto',
              height: img.height ? `${img.height}px` : 'auto',
            }}
          />
        );
      })}

      {/* Charts — placeholder */}
      {charts && charts.map((chart, idx) => {
        if (!rowPositions || !colPositions) return null;
        const top = (rowPositions[chart.position?.row || 0] || 0) - scrollTop;
        const left = (colPositions[chart.position?.col || 0] || 0) - scrollLeft;

        return (
          <div
            key={`chart-${idx}`}
            style={{
              position: 'absolute',
              top: `${top}px`,
              left: `${left}px`,
              width: '300px',
              height: '200px',
              background: 'rgba(102, 126, 234, 0.08)',
              border: '1px dashed rgba(102, 126, 234, 0.3)',
              borderRadius: '8px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: '#667eea',
              fontSize: '12px',
              fontWeight: 500,
              pointerEvents: 'auto',
            }}
          >
            📊 {chart.type || 'Chart'} ({chart.dataRange || 'N/A'})
          </div>
        );
      })}
    </div>
  );
}
