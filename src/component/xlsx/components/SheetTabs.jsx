import React from 'react';

/**
 * SheetTabs — Tab bar for switching between spreadsheet sheets.
 */
export default function SheetTabs({ sheets, activeIndex, onSheetChange }) {
  if (!sheets || sheets.length <= 1) return null;

  return (
    <div className="sv-tabs" role="tablist">
      {sheets.map((sheet, idx) => (
        <div
          key={sheet.name || idx}
          className={`sv-tab ${idx === activeIndex ? 'sv-tab-active' : ''}`}
          role="tab"
          aria-selected={idx === activeIndex}
          onClick={() => onSheetChange(idx)}
        >
          {sheet.name || `Sheet${idx + 1}`}
        </div>
      ))}
    </div>
  );
}
