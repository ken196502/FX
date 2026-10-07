"""
web_shot.py - BOCHK 汇率页面截图
=================================

抓取汇率的同时，用浏览器打开 BOCHK 三个牌价页面并整页截图，
作为邮件附件发送，便于人工核对「当时页面上显示的数字」。

截图目标为 BOCHK iframe 源页面（与 web_rates.py 用 requests 抓取的
URL 完全一致），页面上含「資料更新於香港時間」，可直接对照。

浏览器选择：优先系统 Chrome，其次系统 Edge，最后才用 Playwright
自带的 chromium（channel="chrome" / "msedge" / 默认），避免强依赖
`playwright install` 下载的浏览器。

截图失败（浏览器不可用/页面打不开）只记录警告并返回空列表，
不影响汇率抓取和邮件发送主流程。
"""

from __future__ import annotations

import datetime as dt
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# 截图目标：(中文标签, URL, 文件名前缀)
SHOT_TARGETS: tuple[tuple[str, str, str], ...] = (
    (
        "中銀香港 — 各類貨幣兌港元電匯牌價",
        "https://www.bochk.com/whk/rates/exchangeRatesHKD/"
        "exchangeRatesHKD-input.action?lang=hk",
        "bochk_hkd_rates",
    ),
    (
        "中銀香港 — 各類貨幣兌美元電匯牌價",
        "https://www.bochk.com/whk/rates/exchangeRatesUSD/"
        "exchangeRatesUSD-input.action?lang=hk",
        "bochk_usd_rates",
    ),
    (
        "中銀香港 — 各類貨幣兌港元現鈔牌價",
        "https://www.bochk.com/whk/rates/exchangeRatesForCurrency/"
        "exchangeRatesForCurrency-input.action?lang=hk",
        "bochk_fx_rates",
    ),
)

# 浏览器启动候选：依次尝试，第一个可用的即可
# 优先系统 Chrome，其次 Edge，最后才用 playwright 自带的 chromium
_BROWSER_CHANNELS: tuple[str | None, ...] = ("chrome", "msedge", None)

_VIEWPORT = {"width": 1280, "height": 900}
_TIMEOUT_MS = 60_000


def _launch_browser(pw):
    """依次尝试启动浏览器，返回第一个可用的 Browser 实例。

    Args:
        pw: sync_playwright 启动器对象

    Returns:
        Browser 实例，全部失败时返回 None
    """
    last_exc: Exception | None = None
    for channel in _BROWSER_CHANNELS:
        try:
            if channel:
                browser = pw.chromium.launch(channel=channel)
            else:
                browser = pw.chromium.launch()
            logger.info("[SHOT] 浏览器启动成功: %s", channel or "chromium(playwright)")
            return browser
        except Exception as exc:
            last_exc = exc
            logger.warning("[SHOT] 浏览器启动失败 (%s): %s", channel or "chromium", exc)
    if last_exc is not None:
        logger.warning("[SHOT] 所有浏览器候选均不可用，跳过截图")
    return None


def capture_rate_pages(
    date_str: str,
    output_dir: str | Path,
) -> list[dict]:
    """截图 BOCHK 三个牌价页面。

    Args:
        date_str: 日期字符串 YYYYMMDD，用于文件名
        output_dir: 截图保存目录

    Returns:
        [{"label": 中文标签, "url": 页面 URL, "path": Path}] 列表；
        浏览器不可用或全部失败时返回空列表
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.warning("[SHOT] 未安装 playwright，跳过页面截图")
        return []

    results: list[dict] = []
    try:
        with sync_playwright() as pw:
            browser = _launch_browser(pw)
            if browser is None:
                return []
            try:
                for label, url, prefix in SHOT_TARGETS:
                    path = out_dir / f"{prefix}_{date_str}.png"
                    try:
                        page = browser.new_page(viewport=_VIEWPORT)
                        try:
                            page.goto(
                                url, wait_until="networkidle",
                                timeout=_TIMEOUT_MS,
                            )
                            # 等待汇率表格渲染完成
                            page.wait_for_selector(
                                "table", timeout=_TIMEOUT_MS,
                            )
                            page.wait_for_timeout(500)
                            page.screenshot(path=str(path), full_page=True)
                        finally:
                            page.close()
                    except Exception as exc:
                        logger.warning(
                            "[SHOT] 截图失败 (%s): %s", label, exc,
                        )
                        continue

                    if not path.exists() or path.stat().st_size == 0:
                        logger.warning("[SHOT] 截图为空 (%s)", label)
                        continue

                    size_kb = path.stat().st_size / 1024
                    logger.info(
                        "[SHOT] %s -> %s (%.0f KB)", label, path.name, size_kb,
                    )
                    results.append({
                        "label": label,
                        "url": url,
                        "path": path,
                    })
            finally:
                browser.close()
    except Exception as exc:
        logger.warning("[SHOT] 截图流程异常，已跳过: %s", exc)
        return results

    return results


def capture_with_timestamp(
    date_str: str | None = None,
    output_dir: str | Path = "temp",
) -> list[dict]:
    """便捷入口：用当前日期截图到指定目录。

    Args:
        date_str: 日期字符串 YYYYMMDD，None 表示今天
        output_dir: 截图保存目录

    Returns:
        同 capture_rate_pages
    """
    if date_str is None:
        date_str = dt.date.today().strftime("%Y%m%d")
    return capture_rate_pages(date_str, output_dir)
