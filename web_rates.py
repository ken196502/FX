"""
web_rates.py - 从网站抓取银行汇率和交易所印花税率
===================================================

数据来源：
  1. HKEx 港交所「用於計算印花稅的匯率」
     https://www.hkex.com.hk/chi/market/sec_tradinfo/stampfx/stampfx_c.asp
     — 日历页面，每个工作日链接到一个 .xls 文件，
       xls 内含人民币和美元兑港元的汇率（用于计算印花税）。

  2. BOCHK 中銀香港「各類貨幣兌港元電匯牌價」
     https://www.bochk.com/tc/investment/rates/hkdrates.html
     — iframe 加载，含各币种兑港元的电汇买入/卖出价。

  3. BOCHK 中銀香港「各類貨幣兌港元現鈔牌價」
     https://www.bochk.com/tc/investment/rates/fxrates.html
     — iframe 加载，含各币种兑港元的现钞买入/卖出价。
"""

from __future__ import annotations

import datetime as dt
import logging
import re
from io import BytesIO
from pathlib import Path

import requests
import xlrd

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 通用配置
# ---------------------------------------------------------------------------

_TIMEOUT = 30
_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
}

HKEX_STAMPFX_URL = (
    "https://www.hkex.com.hk/chi/market/sec_tradinfo/stampfx/stampfx_c.asp"
)
HKEX_STAMPFX_BASE = "https://www.hkex.com.hk"

BOCHK_HKDRATES_IFRAME = (
    "https://www.bochk.com/whk/rates/exchangeRatesHKD/"
    "exchangeRatesHKD-input.action?lang=hk"
)
BOCHK_FXRATES_IFRAME = (
    "https://www.bochk.com/whk/rates/exchangeRatesForCurrency/"
    "exchangeRatesForCurrency-input.action?lang=hk"
)
BOCHK_USDRATES_IFRAME = (
    "https://www.bochk.com/whk/rates/exchangeRatesUSD/"
    "exchangeRatesUSD-input.action?lang=hk"
)

# 面向用户的页面 URL（用于邮件正文展示来源链接）
BOCHK_HKDRATES_PAGE = "https://www.bochk.com/tc/investment/rates/hkdrates.html"
BOCHK_USDRATES_PAGE = "https://www.bochk.com/tc/investment/rates/usdrates.html"
BOCHK_FXRATES_PAGE = "https://www.bochk.com/tc/investment/rates/fxrates.html"


# ---------------------------------------------------------------------------
# HKEx 港交所印花税率
# ---------------------------------------------------------------------------

# 匹配日历中的 xls 链接和对应日期
_HKEX_LINK_RE = re.compile(
    r'href=([^>]+/(\d{8})\.xls)>(\d+)</a>',
    re.IGNORECASE,
)


def fetch_hkex_stampfx(date_str: str | None = None) -> dict:
    """抓取 HKEx 印花税率页面，下载指定日期的 xls 并解析汇率。

    页面以日历形式展示最近两个月的工作日，每天链接到一个 .xls 文件，
    xls 内含人民币（Renminbi）和美元（U.S. dollars）兑港元的汇率。

    Args:
        date_str: YYYYMMDD 格式日期，None 表示今天。
            如果指定日期在日历中不存在（如周末/假日），
            将查找日历中最近的之前的工作日。

    Returns:
        包含以下键的字典：
          - date: 日期字符串 YYYYMMDD
          - rates: 汇率列表，每项含 currency / unit / hkd_rate
          - xls_url: xls 文件 URL
          - update_time: 资料更新时间（如有）
        如果未找到数据返回空字典。
    """
    if date_str is None:
        date_str = dt.date.today().strftime("%Y%m%d")

    logger.info("[HKEX] 抓取印花税率页面, 目标日期=%s", date_str)

    resp = requests.get(
        HKEX_STAMPFX_URL, headers=_HEADERS, timeout=_TIMEOUT
    )
    resp.encoding = "big5"
    resp.raise_for_status()
    html = resp.text

    # 提取所有 (url, yyyymmdd, day) 匹配
    links: list[tuple[str, str]] = []
    for m in _HKEX_LINK_RE.finditer(html):
        url = m.group(1).strip()
        yyyymmdd = m.group(2)
        links.append((url, yyyymmdd))

    if not links:
        logger.warning("[HKEX] 页面中未找到任何 xls 链接")
        return {}

    # 找到目标日期或之前最近的工作日
    target = date_str
    matching = [lk for lk in links if lk[1] == target]
    if not matching:
        # 查找 <= target 的最大日期
        earlier = [lk for lk in links if lk[1] <= target]
        if earlier:
            earlier.sort(key=lambda x: x[1], reverse=True)
            target = earlier[0][1]
            matching = [earlier[0]]
            logger.info("[HKEX] 目标日期 %s 无数据, 使用最近工作日 %s", date_str, target)
        else:
            logger.warning("[HKEX] 日历中无 <= %s 的日期", target)
            return {}

    xls_path = matching[0][0]
    if xls_path.startswith("/"):
        xls_url = HKEX_STAMPFX_BASE + xls_path
    else:
        xls_url = xls_path

    logger.info("[HKEX] 下载 xls: %s", xls_url)
    xls_resp = requests.get(xls_url, headers=_HEADERS, timeout=_TIMEOUT)
    xls_resp.raise_for_status()

    rates = _parse_hkex_xls(xls_resp.content)
    if not rates:
        logger.warning("[HKEX] xls 解析无结果: %s", xls_url)
        return {}

    return {
        "date": target,
        "rates": rates,
        "xls_url": xls_url,
        "update_time": "",
    }


def _parse_hkex_xls(content: bytes) -> list[dict]:
    """解析 HKEx stampfx xls 文件内容。

    xls 结构：
      row 0: 说明文字
      row 4: Date | 日期序列号 | (空)
      row 5: 货币名称 | 单位 | 港元等值(Buying T.T.)
      row 6: Renminbi | 1 | 1.1565
      row 7: U.S. dollars | 1 | 7.8145

    Args:
        content: xls 文件的原始字节内容

    Returns:
        汇率字典列表，每项含:
          - currency: 货币名称（如 "Renminbi" / "U.S. dollars"）
          - unit: 单位（通常为 1）
          - hkd_rate: 港元汇率
    """
    wb = xlrd.open_workbook(file_contents=content)
    ws = wb.sheet_by_index(0)

    rates: list[dict] = []
    # 从 row 6 开始读取数据行（row 5 是表头）
    for row_idx in range(6, ws.nrows):
        name = ws.cell_value(row_idx, 0)
        if not name:
            continue
        # 货币名称可能含中英文换行，取第一行英文部分
        currency = str(name).split("\n")[0].strip()
        unit = ws.cell_value(row_idx, 1)
        hkd_rate = ws.cell_value(row_idx, 2)

        rates.append({
            "currency": currency,
            "unit": float(unit) if unit else 1.0,
            "hkd_rate": float(hkd_rate) if hkd_rate else 0.0,
        })

    wb.release_resources()
    return rates


# ---------------------------------------------------------------------------
# BOCHK 中銀香港 — 港币电汇牌价 / 现钞牌价
# ---------------------------------------------------------------------------

# 匹配 <tr> 中的三个 <td> 内容
_BOCHK_TD_RE = re.compile(
    r'<td[^>]*>(.*?)</td>',
    re.DOTALL,
)

# 匹配表格中的数字（含小数）
_NUMBER_RE = re.compile(r'[\d.,]+')


def fetch_bochk_hkdrates() -> dict:
    """抓取 BOCHK 各類貨幣兌港元電匯牌價。

    通过 iframe URL 直接获取 HTML，解析 table 中的
    货币名称、客户卖出价、客户买入价。

    Returns:
        包含以下键的字典：
          - rates: 汇率列表，每项含 currency / sell / buy
          - update_time: 资料更新时间字符串
    """
    logger.info("[BOCHK-HKD] 抓取港币电汇牌价")
    html = _fetch_bochk_iframe(BOCHK_HKDRATES_IFRAME)
    rates, update_time = _parse_bochk_rates_table(html)
    return {"rates": rates, "update_time": update_time}


def fetch_bochk_fxrates() -> dict:
    """抓取 BOCHK 各類貨幣兌港元現鈔牌價。

    Returns:
        包含以下键的字典：
          - rates: 汇率列表，每项含 currency / sell / buy
          - update_time: 资料更新时间字符串
    """
    logger.info("[BOCHK-FX] 抓取港币现钞牌价")
    html = _fetch_bochk_iframe(BOCHK_FXRATES_IFRAME)
    rates, update_time = _parse_bochk_rates_table(html)
    return {"rates": rates, "update_time": update_time}


def fetch_bochk_usdrates() -> dict:
    """抓取 BOCHK 各類貨幣兌美元電匯牌價。

    页面提供各币种兑美元的汇率，货币名格式为 "美元/人民幣"、
    "美元/港元" 等，部分为反向如 "澳元/美元"、"英鎊/美元"。

    Returns:
        包含以下键的字典：
          - rates: 汇率列表，每项含 currency / sell / buy
          - update_time: 资料更新时间字符串
    """
    logger.info("[BOCHK-USD] 抓取美元电汇牌价")
    html = _fetch_bochk_iframe(
        BOCHK_USDRATES_IFRAME,
        referer=BOCHK_USDRATES_PAGE,
    )
    rates, update_time = _parse_bochk_rates_table(html)
    return {"rates": rates, "update_time": update_time}


def _fetch_bochk_iframe(url: str, referer: str = "") -> str:
    """获取 BOCHK iframe 页面 HTML 内容。

    Args:
        url: iframe 的 URL
        referer: Referer 页面 URL，默认用 HKD 牌价页

    Returns:
        HTML 文本
    """
    headers = {
        **_HEADERS,
        "Referer": referer or BOCHK_HKDRATES_PAGE,
    }
    resp = requests.get(url, headers=headers, timeout=_TIMEOUT)
    resp.encoding = "utf-8"
    resp.raise_for_status()
    return resp.text


def _parse_bochk_rates_table(html: str) -> tuple[list[dict], str]:
    """解析 BOCHK 汇率表格 HTML。

    表格结构：
      <table class="form_table import-data second-right">
        <tr><th>貨幣</th><th>客戶賣出</th><th>客戶買入</th></tr>
        <tr><td>人民幣(在岸)</td><td>1.162780</td><td>1.175420</td></tr>
        ...
      </table>
      <table class="form_table">
        <tr><td><b>資料更新於香港時間： 2026/09/29 13:42:59</b></td></tr>
      </table>

    Args:
        html: iframe 页面的 HTML 文本

    Returns:
        (汇率列表, 更新时间字符串)
        汇率列表每项含 currency / sell / buy
    """
    rates: list[dict] = []
    update_time = ""

    # 提取更新时间
    time_match = re.search(
        r'資料更新於香港時間[：:]\s*([\d/ :]+)',
        html,
    )
    if time_match:
        update_time = time_match.group(1).strip()

    # 找到 form_table import-data 表格区域
    # 按 <tr> 分割，找到含三个 <td> 且后两个为数字的行
    tr_re = re.compile(r'<tr[^>]*>(.*?)</tr>', re.DOTALL)
    for tr_match in tr_re.finditer(html):
        tr_content = tr_match.group(1)
        tds = _BOCHK_TD_RE.findall(tr_content)
        if len(tds) < 3:
            continue

        currency = _clean_html_text(tds[0])
        sell_str = _clean_html_text(tds[1])
        buy_str = _clean_html_text(tds[2])

        # 跳过表头行
        if "貨幣" in currency or "客户" in sell_str or "客戶" in sell_str:
            continue

        # 跳过空行
        if not currency:
            continue

        sell = _extract_number(sell_str)
        buy = _extract_number(buy_str)

        if sell is None and buy is None:
            continue

        rates.append({
            "currency": currency,
            "sell": sell,
            "buy": buy,
        })

    return rates, update_time


def _clean_html_text(text: str) -> str:
    """清理 HTML 标签和多余空白。

    Args:
        text: 含 HTML 标签的原始文本

    Returns:
        清理后的纯文本
    """
    # 去掉所有 HTML 标签
    cleaned = re.sub(r'<[^>]+>', '', text)
    # 替换 &nbsp; 和多余空白
    cleaned = cleaned.replace('\xa0', ' ')
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    return cleaned


def _extract_number(text: str) -> float | None:
    """从文本中提取第一个数字（支持小数和千分位逗号）。

    Args:
        text: 可能含数字的文本

    Returns:
        float 或 None
    """
    m = _NUMBER_RE.search(text)
    if not m:
        return None
    num_str = m.group().replace(',', '')
    return float(num_str)


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------

def _print_hkex(data: dict) -> None:
    """打印 HKEx 印花税率结果。"""
    if not data:
        print("[HKEx] 未获取到印花税率数据")
        return

    print(f"\n[HKEx] 用於計算印花稅的匯率 — {data['date']}")
    print(f"  来源: {data['xls_url']}")
    for r in data["rates"]:
        print(
            f"  {r['currency']}: "
            f"1 {r['currency']} = {r['hkd_rate']} HKD"
        )


def _print_bochk(title: str, data: dict) -> None:
    """打印 BOCHK 汇率结果。"""
    rates = data.get("rates", [])
    update_time = data.get("update_time", "")

    print(f"\n[{title}]")
    if update_time:
        print(f"  资料更新于香港时间: {update_time}")
    if not rates:
        print("  未获取到汇率数据")
        return

    print(f"  共 {len(rates)} 种货币:")
    for r in rates:
        sell = f"{r['sell']:.6f}" if r['sell'] is not None else "N/A"
        buy = f"{r['buy']:.6f}" if r['buy'] is not None else "N/A"
        print(f"    {r['currency']}: 客户卖出={sell}, 客户买入={buy}")


# ---------------------------------------------------------------------------
# 主入口
# ---------------------------------------------------------------------------

def process_web_rates(
    date_str: str | None = None,
    output_dir: str | None = None,
) -> None:
    """抓取银行和交易所汇率并打印。

    依次抓取：
      1. HKEx 港交所印花税率
      2. BOCHK 港币电汇牌价
      3. BOCHK 港币现钞牌价

    如指定 output_dir，同时将结果保存为文本文件。

    Args:
        date_str: YYYYMMDD 格式日期，None 表示今天（仅影响 HKEx）
        output_dir: 输出目录，None 则不保存文件
    """
    if date_str is None:
        date_str = dt.date.today().strftime("%Y%m%d")

    logger.info("[WEB] 开始抓取汇率, date=%s", date_str)

    # 1. HKEx 印花税率
    hkex_data = fetch_hkex_stampfx(date_str)
    _print_hkex(hkex_data)

    # 2. BOCHK 港币电汇牌价
    bochk_hkd = fetch_bochk_hkdrates()
    _print_bochk("BOCHK 港币电汇牌价", bochk_hkd)

    # 3. BOCHK 港币现钞牌价
    bochk_fx = fetch_bochk_fxrates()
    _print_bochk("BOCHK 港币现钞牌价", bochk_fx)

    # 保存到文件
    if output_dir:
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"web_rates_{date_str}.txt"
        _save_to_file(out_path, hkex_data, bochk_hkd, bochk_fx)
        print(f"\n[WEB] 结果已保存到: {out_path}")

    logger.info("[WEB] 抓取完成")


def _save_to_file(
    path: Path,
    hkex_data: dict,
    bochk_hkd: dict,
    bochk_fx: dict,
) -> None:
    """将抓取结果保存到文本文件。

    Args:
        path: 输出文件路径
        hkex_data: HKEx 数据
        bochk_hkd: BOCHK 电汇牌价数据
        bochk_fx: BOCHK 现钞牌价数据
    """
    lines: list[str] = []

    lines.append(f"汇率抓取结果 — {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")

    # HKEx
    if hkex_data:
        lines.append(f"[HKEx] 用於計算印花稅的匯率 — {hkex_data['date']}")
        lines.append(f"  来源: {hkex_data['xls_url']}")
        for r in hkex_data["rates"]:
            lines.append(
                f"  {r['currency']}: 1 {r['currency']} = {r['hkd_rate']} HKD"
            )
    else:
        lines.append("[HKEx] 未获取到印花税率数据")
    lines.append("")

    # BOCHK HKD
    lines.append(f"[BOCHK 港币电汇牌价]")
    if bochk_hkd.get("update_time"):
        lines.append(f"  资料更新于香港时间: {bochk_hkd['update_time']}")
    for r in bochk_hkd.get("rates", []):
        sell = f"{r['sell']:.6f}" if r['sell'] is not None else "N/A"
        buy = f"{r['buy']:.6f}" if r['buy'] is not None else "N/A"
        lines.append(f"  {r['currency']}: 客户卖出={sell}, 客户买入={buy}")
    lines.append("")

    # BOCHK FX
    lines.append(f"[BOCHK 港币现钞牌价]")
    if bochk_fx.get("update_time"):
        lines.append(f"  资料更新于香港时间: {bochk_fx['update_time']}")
    for r in bochk_fx.get("rates", []):
        sell = f"{r['sell']:.6f}" if r['sell'] is not None else "N/A"
        buy = f"{r['buy']:.6f}" if r['buy'] is not None else "N/A"
        lines.append(f"  {r['currency']}: 客户卖出={sell}, 客户买入={buy}")

    path.write_text("\n".join(lines), encoding="utf-8")
