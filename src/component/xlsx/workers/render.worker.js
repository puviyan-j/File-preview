/**
 * Render Worker (Stub)
 *
 * OffscreenCanvas rendering worker. Falls back to main-thread rendering
 * when OffscreenCanvas is not supported.
 *
 * For the initial implementation, rendering is done on the main thread
 * via the renderEngine module. This worker can be progressively enhanced
 * to offload rendering to a separate thread.
 */

let canvas = null;
let ctx = null;

self.onmessage = function (e) {
  const { type } = e.data;

  switch (type) {
    case 'INIT_CANVAS': {
      canvas = e.data.canvas;
      ctx = canvas.getContext('2d');
      self.postMessage({ type: 'READY' });
      break;
    }

    case 'RENDER_TILE': {
      // Future: receive tile spec + data, render on OffscreenCanvas
      // For now, post back acknowledgement
      self.postMessage({
        type: 'TILE_RENDERED',
        key: e.data.key,
      });
      break;
    }

    case 'RESIZE': {
      if (canvas) {
        canvas.width = e.data.width;
        canvas.height = e.data.height;
      }
      break;
    }
  }
};
