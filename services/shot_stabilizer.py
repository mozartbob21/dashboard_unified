"""Стабилизация страницы перед скриншотом.
Убивает анимации/перерисовки, чтобы каждый снимок был одинаково полным."""

KILL_CSS = """
*, *::before, *::after {
  animation-duration: 0s !important;
  animation-delay: 0s !important;
  animation-iteration-count: 1 !important;
  transition-duration: 0s !important;
  transition-delay: 0s !important;
  caret-color: transparent !important;
}
html, body { scroll-behavior: auto !important; }
"""

_JS = """() => {
    const A = document.getAnimations ? document.getAnimations() : [];
    for (const a of A) {
        try { a.finish(); } catch (e) { try { a.cancel(); } catch (e2) {} }
    }
    if (document.fonts && document.fonts.ready) { document.fonts.ready.catch(() => {}); }
}"""


def stabilize_page(page, settle_ms=400):
    """Sync Playwright. Вызывать ПЕРЕД page.screenshot(...)."""
    try:
        page.emulate_media(reduced_motion="reduce")
    except Exception:
        pass
    try:
        page.add_style_tag(content=KILL_CSS)
    except Exception:
        pass
    try:
        page.evaluate(_JS)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=3000)
    except Exception:
        pass
    if settle_ms:
        page.wait_for_timeout(settle_ms)


async def stabilize_page_async(page, settle_ms=400):
    """Async Playwright. Вызывать: await stabilize_page_async(page)."""
    try:
        await page.emulate_media(reduced_motion="reduce")
    except Exception:
        pass
    try:
        await page.add_style_tag(content=KILL_CSS)
    except Exception:
        pass
    try:
        await page.evaluate(_JS)
    except Exception:
        pass
    try:
        await page.wait_for_load_state("networkidle", timeout=3000)
    except Exception:
        pass
    if settle_ms:
        await page.wait_for_timeout(settle_ms)
