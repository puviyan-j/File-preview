/**
 * Scroll Engine
 *
 * Provides smooth, 60fps scrolling with momentum support.
 * Encapsulates scroll state and coordinates with rAF rendering loop.
 */

export function createScrollEngine(options = {}) {
  const {
    onScroll = () => {},
    friction = 0.92,
    minVelocity = 0.5,
  } = options;

  let scrollTop = 0;
  let scrollLeft = 0;
  let maxScrollTop = 0;
  let maxScrollLeft = 0;

  // Momentum state
  let velocityX = 0;
  let velocityY = 0;
  let momentumActive = false;
  let rafId = null;

  // Throttle: last scroll event timestamp
  let lastScrollTime = 0;

  function setContentSize(contentWidth, contentHeight, viewportWidth, viewportHeight) {
    maxScrollTop = Math.max(0, contentHeight - viewportHeight);
    maxScrollLeft = Math.max(0, contentWidth - viewportWidth);
    // Clamp current position
    scrollTop = Math.min(scrollTop, maxScrollTop);
    scrollLeft = Math.min(scrollLeft, maxScrollLeft);
  }

  function handleWheel(e) {
    e.preventDefault();

    // Use deltaX/deltaY directly for precise scroll
    let dx = e.deltaX;
    let dy = e.deltaY;

    // Handle deltaMode (0 = pixels, 1 = lines, 2 = pages)
    if (e.deltaMode === 1) {
      dx *= 24;
      dy *= 24;
    } else if (e.deltaMode === 2) {
      dx *= 400;
      dy *= 400;
    }

    // Apply scroll
    scrollLeft = clamp(scrollLeft + dx, 0, maxScrollLeft);
    scrollTop = clamp(scrollTop + dy, 0, maxScrollTop);

    // Record velocity for momentum
    velocityX = dx;
    velocityY = dy;

    // Notify
    const now = performance.now();
    if (now - lastScrollTime > 8) { // ~120fps throttle
      lastScrollTime = now;
      onScroll(scrollTop, scrollLeft);
    }

    // Start momentum when wheel stops
    stopMomentum();
    // Small delay to detect wheel stop
  }

  function startMomentum() {
    if (momentumActive) return;
    if (Math.abs(velocityX) < minVelocity && Math.abs(velocityY) < minVelocity) return;

    momentumActive = true;
    momentumLoop();
  }

  function momentumLoop() {
    if (!momentumActive) return;

    velocityX *= friction;
    velocityY *= friction;

    if (Math.abs(velocityX) < minVelocity && Math.abs(velocityY) < minVelocity) {
      momentumActive = false;
      return;
    }

    scrollLeft = clamp(scrollLeft + velocityX, 0, maxScrollLeft);
    scrollTop = clamp(scrollTop + velocityY, 0, maxScrollTop);
    onScroll(scrollTop, scrollLeft);

    rafId = requestAnimationFrame(momentumLoop);
  }

  function stopMomentum() {
    momentumActive = false;
    if (rafId) {
      cancelAnimationFrame(rafId);
      rafId = null;
    }
  }

  function scrollTo(top, left) {
    stopMomentum();
    scrollTop = clamp(top, 0, maxScrollTop);
    scrollLeft = clamp(left, 0, maxScrollLeft);
    onScroll(scrollTop, scrollLeft);
  }

  function getPosition() {
    return { scrollTop, scrollLeft };
  }

  function destroy() {
    stopMomentum();
  }

  return {
    handleWheel,
    startMomentum,
    stopMomentum,
    scrollTo,
    setContentSize,
    getPosition,
    destroy,
    get scrollTop() { return scrollTop; },
    get scrollLeft() { return scrollLeft; },
  };
}

function clamp(val, min, max) {
  return Math.max(min, Math.min(max, val));
}
