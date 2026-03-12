/**
 * Tile Manager
 *
 * Divides the canvas into 512×512px tiles, manages an LRU cache
 * of rendered tiles, and coordinates rendering requests.
 */

const TILE_SIZE = 512;
const MAX_TILES = 64; // LRU cache limit

/**
 * Create a tile manager for a sheet.
 */
export function createTileManager() {
  const cache = new Map(); // tileKey → { bitmap, rowStart, rowEnd, colStart, colEnd, dirty }
  const lruOrder = [];     // oldest → newest tile keys

  function tileKey(tileRow, tileCol) {
    return `${tileRow}:${tileCol}`;
  }

  /**
   * Given a visible range and position arrays, compute which tile grid coordinates
   * are needed and return the tile specs.
   */
  function getTilesForRange(visibleRange, rowPositions, colPositions) {
    const { startRow, endRow, startCol, endCol } = visibleRange;
    if (!rowPositions || !colPositions) return [];

    const topPx = rowPositions[startRow] || 0;
    const bottomPx = rowPositions[endRow] || rowPositions[rowPositions.length - 1] || 0;
    const leftPx = colPositions[startCol] || 0;
    const rightPx = colPositions[endCol] || colPositions[colPositions.length - 1] || 0;

    const tileRowStart = Math.floor(topPx / TILE_SIZE);
    const tileRowEnd = Math.ceil(bottomPx / TILE_SIZE);
    const tileColStart = Math.floor(leftPx / TILE_SIZE);
    const tileColEnd = Math.ceil(rightPx / TILE_SIZE);

    const tiles = [];
    for (let tr = tileRowStart; tr < tileRowEnd; tr++) {
      for (let tc = tileColStart; tc < tileColEnd; tc++) {
        const key = tileKey(tr, tc);
        tiles.push({
          key,
          tileRow: tr,
          tileCol: tc,
          x: tc * TILE_SIZE,
          y: tr * TILE_SIZE,
          width: TILE_SIZE,
          height: TILE_SIZE,
          cached: cache.has(key) && !cache.get(key).dirty,
        });
      }
    }

    return tiles;
  }

  /**
   * Store a rendered tile bitmap.
   */
  function setTile(key, bitmap) {
    // Evict if at capacity
    while (cache.size >= MAX_TILES && lruOrder.length > 0) {
      const evictKey = lruOrder.shift();
      cache.delete(evictKey);
    }

    cache.set(key, { bitmap, dirty: false });

    // Move to end of LRU
    const idx = lruOrder.indexOf(key);
    if (idx >= 0) lruOrder.splice(idx, 1);
    lruOrder.push(key);
  }

  /**
   * Get a cached tile bitmap.
   */
  function getTile(key) {
    const entry = cache.get(key);
    if (!entry || entry.dirty) return null;

    // Touch LRU
    const idx = lruOrder.indexOf(key);
    if (idx >= 0) lruOrder.splice(idx, 1);
    lruOrder.push(key);

    return entry.bitmap;
  }

  /**
   * Invalidate tiles that overlap the given row range (when new data arrives).
   */
  function invalidateRows(startRow, endRow, rowPositions) {
    if (!rowPositions) {
      // Invalidate everything
      for (const [, entry] of cache) entry.dirty = true;
      return;
    }

    const topPx = rowPositions[startRow] || 0;
    const bottomPx = rowPositions[Math.min(endRow, rowPositions.length - 1)] || 0;
    const tileRowStart = Math.floor(topPx / TILE_SIZE);
    const tileRowEnd = Math.ceil(bottomPx / TILE_SIZE);

    for (const [key, entry] of cache) {
      const [tr] = key.split(':').map(Number);
      if (tr >= tileRowStart && tr < tileRowEnd) {
        entry.dirty = true;
      }
    }
  }

  /**
   * Invalidate all tiles.
   */
  function invalidateAll() {
    for (const [, entry] of cache) entry.dirty = true;
  }

  /**
   * Clear the entire cache.
   */
  function clear() {
    cache.clear();
    lruOrder.length = 0;
  }

  return {
    getTilesForRange,
    setTile,
    getTile,
    invalidateRows,
    invalidateAll,
    clear,
    get size() { return cache.size; },
    TILE_SIZE,
  };
}
