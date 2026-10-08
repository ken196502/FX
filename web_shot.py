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
from pathlib import Path

from log_setup import get_logger

logger = get_logger(__name__)

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


def _shot_time(path: Path) -> str:
    """取截图的抓取时间（文件修改时间）。

    Args:
        path: 截图文件路径

    Returns:
        "YYYY-MM-DD HH:MM" 格式时间字符串
    """
    return dt.datetime.fromtimestamp(path.stat().st_mtime).strftime(
        "%Y-%m-%d %H:%M"
    )


def _shot_path(output_dir: Path, prefix: str, date_str: str) -> Path:
    """按统一命名规则生成截图文件路径。

    Args:
        output_dir: 截图保存目录
        prefix: 文件名前缀（SHOT_TARGETS 第三项）
        date_str: 日期字符串 YYYYMMDD

    Returns:
        截图文件路径
    """
    return Path(output_dir) / f"{prefix}_{date_str}.png"


def existing_rate_page_shots(
    date_str: str,
    output_dir: str | Path,
) -> list[dict]:
    """返回当日已存在的 BOCHK 页面截图（不启动浏览器）。

    用于同日多次运行时复用已生成的截图，避免重复打开浏览器截图。

    Args:
        date_str: 日期字符串 YYYYMMDD，用于文件名
        output_dir: 截图保存目录

    Returns:
        [{"label": 中文标签, "url": 页面 URL, "path": Path,
          "captured_at": 截图时间 "YYYY-MM-DD HH:MM"}] 列表；
        文件不存在或为空时该项不返回
    """
    out_dir = Path(output_dir)
    results: list[dict] = []
    for label, url, prefix in SHOT_TARGETS:
        path = _shot_path(out_dir, prefix, date_str)
        if path.exists() and path.stat().st_size > 0:
            results.append({
                "label": label,
                "url": url,
                "path": path,
                "captured_at": _shot_time(path),
            })
    if results:
        logger.info(
            "[SHOT] 复用当日已有截图 %d 张: %s",
            len(results), ", ".join(str(r["path"].name) for r in results),
        )
    return results


def capture_rate_pages(
    date_str: str,
    output_dir: str | Path,
    targets: tuple[tuple[str, str, str], ...] = SHOT_TARGETS,
) -> list[dict]:
    """截图 BOCHK 牌价页面。

    Args:
        date_str: 日期字符串 YYYYMMDD，用于文件名
        output_dir: 截图保存目录
        targets: 需要截图的页面列表，默认全部三个牌价页。
            传入子集时只截缺失的页面，已有截图保持不动。

    Returns:
        [{"label": 中文标签, "url": 页面 URL, "path": Path,
          "captured_at": 截图时间 "YYYY-MM-DD HH:MM"}] 列表；
        浏览器不可用或全部失败时返回空列表
    """
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        logger.warning("[SHOT] 未安装 playwright，跳过页面截图")
        return []

    started_at = dt.datetime.now()
    logger.info(
        "[SHOT] 开始截图 %d 个页面: %s",
        len(targets), ", ".join(t[2] for t in targets),
    )

    results: list[dict] = []
    try:
        with sync_playwright() as pw:
            browser = _launch_browser(pw)
            if browser is None:
                return []
            try:
                for label, url, prefix in targets:
                    path = _shot_path(out_dir, prefix, date_str)
                    page_started_at = dt.datetime.now()
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
                        "[SHOT] %s -> %s (%.0f KB, 耗时 %.1fs)",
                        label, path.name, size_kb,
                        (dt.datetime.now() - page_started_at).total_seconds(),
                    )
                    results.append({
                        "label": label,
                        "url": url,
                        "path": path,
                        "captured_at": _shot_time(path),
                    })
            finally:
                browser.close()
    except Exception as exc:
        logger.warning("[SHOT] 截图流程异常，已跳过: %s", exc)
        return results

    logger.info(
        "[SHOT] 截图完成: 成功 %d/%d 张, 耗时 %.1fs",
        len(results), len(targets),
        (dt.datetime.now() - started_at).total_seconds(),
    )
    if len(results) < len(targets):
        logger.warning(
            "[SHOT] 有 %d 个页面未截到图，邮件附件中可能缺少对应截图",
            len(targets) - len(results),
        )

    return results


def resolve_rate_page_shots(
    date_str: str,
    output_dir: str | Path,
    allow_reuse: bool = True,
) -> list[dict]:
    """获取当日 BOCHK 页面截图，优先复用已有截图。

    当日已存在的截图直接复用（不再启动浏览器），只对缺失的页面重新截图，
    便于同日多次运行共用同一批截图。

    Args:
        date_str: 日期字符串 YYYYMMDD
        output_dir: 截图保存目录
        allow_reuse: 是否允许复用当日已有截图，False 则全部重新截图

    Returns:
        截图信息列表，顺序与 SHOT_TARGETS 一致
    """
    out_dir = Path(output_dir)
    shots = existing_rate_page_shots(date_str, out_dir) if allow_reuse else []
    have = {s["path"].name for s in shots}
    missing = tuple(
        t for t in SHOT_TARGETS
        if _shot_path(out_dir, t[2], date_str).name not in have
    )
    logger.info(
        "[SHOT] 截图需求: 复用 %d 张, 待补 %d 张 (allow_reuse=%s)",
        len(shots), len(missing), allow_reuse,
    )
    if missing:
        logger.info(
            "[SHOT] 当日仍有 %d 个页面无截图，开始补截图: %s",
            len(missing), ", ".join(t[2] for t in missing),
        )
        shots.extend(capture_rate_pages(date_str, out_dir, targets=missing))

    order = {
        _shot_path(out_dir, prefix, date_str).name: idx
        for idx, (_, _, prefix) in enumerate(SHOT_TARGETS)
    }
    final = sorted(shots, key=lambda s: order.get(s["path"].name, len(order)))
    logger.info("[SHOT] 当日可用截图共 %d 张", len(final))
    return final


def cleanup_previous_shots(
    output_dir: str | Path,
    keep_date_str: str,
) -> int:
    """删除早于指定日期的 BOCHK 截图。

    当日截图需要保留以便同日后续运行复用，因此只在每次运行开始时
    清理历史日期的 png，避免在 temp 目录无限堆积。

    Args:
        output_dir: 截图保存目录
        keep_date_str: 保留该日期（YYYYMMDD）及其之后的截图

    Returns:
        成功删除的文件数量
    """
    out_dir = Path(output_dir)
    if not out_dir.exists():
        return 0
    prefixes = {prefix for _, _, prefix in SHOT_TARGETS}
    removed = 0
    logger.debug(
        "[SHOT] 开始清理 %s 中早于 %s 的历史截图", out_dir, keep_date_str,
    )
    for path in out_dir.glob("*.png"):
        name = path.stem
        if "_" not in name:
            continue
        prefix, _, date_part = name.rpartition("_")
        if prefix not in prefixes or len(date_part) != 8 or not date_part.isdigit():
            continue
        if date_part >= keep_date_str:
            continue
        try:
            path.unlink()
            removed += 1
            logger.info("[SHOT] 已清理历史日期截图: %s", path.name)
        except Exception as exc:
            logger.warning("[SHOT] 清理截图失败: %s (%s)", path, exc)
    if removed:
        logger.info("[SHOT] 历史截图清理完成, 共删除 %d 个文件", removed)
    return removed


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
