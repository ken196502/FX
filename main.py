"""
main.py - FX 汇率邮件发送入口
==============================

用法：
    uv run fxmail              # 发送今日汇率邮件
    uv run fxmail --date 20260930  # 发送指定日期汇率邮件

流程：
    1. 从 BOCHK / HKEx 网站获取汇率
    2. 按 TFISF Excel 公式计算自定义汇率
    3. 生成两个 Excel 附件（自定义汇率 + 交易所汇率）
    4. 构建邮件正文（含公式和计算结果）
    5. 通过 MS Graph API 发送邮件到 FX_RECIEVER

异常策略：
    正常时只发汇率邮件，不发企业微信 webhook；
    出现任何异常（网页打不开 / 汇率获取不到 / 汇率不是当日）时，
    不发送汇率邮件（也不带汇率附件），只发送企业微信错误通知 +
    一封【错误报告】邮件，并以退出码 1 结束。
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import os
import sys
import traceback
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

from ms_mail import send_mail
from sc_infra import send_error
from rate_export import (
    DEFAULT_MARKDOWN_FACTOR,
    MARKDOWN_FACTORS,
    process_rate_export,
)
from web_rates import (
    BOCHK_HKDRATES_PAGE,
    BOCHK_USDRATES_PAGE,
    HKEX_STAMPFX_URL,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


# -----------------------------------------------------------------------
# 邮件正文构建
# -----------------------------------------------------------------------

def _link(url: str, text: str | None = None) -> str:
    """生成可点击的 HTML 超链接。

    Args:
        url: 链接地址
        text: 显示文本，None 则显示 URL 本身

    Returns:
        HTML <a> 标签字符串
    """
    if not url:
        return ""
    display = text or url
    return (
        f"<a href='{url}' style='color: #1a73e8; "
        f"text-decoration: underline;'>{display}</a>"
    )


def _build_sources_html(result: dict) -> str:
    """构建「数据来源」区块，列出各数据项的来源页面 URL。

    Args:
        result: process_rate_export 返回的结果字典

    Returns:
        HTML 片段字符串
    """
    hkex_data = result.get("hkex_data") or {}
    xls_url = hkex_data.get("xls_url", "")
    hkex_date = hkex_data.get("date", "")

    sources: list[tuple[str, str, str]] = [
        (
            "自定义汇率（F/H，各币种兑港元）",
            BOCHK_HKDRATES_PAGE,
            "中銀香港 — 各類貨幣兌港元電匯牌價",
        ),
        (
            "自定义汇率（交叉汇率，如 USD/CNY）",
            BOCHK_USDRATES_PAGE,
            "中銀香港 — 各類貨幣兌美元電匯牌價",
        ),
        (
            "交易所汇率（印花税率）",
            HKEX_STAMPFX_URL,
            "HKEx — 用於計算印花稅的匯率",
        ),
    ]
    if xls_url:
        label = f"HKEx — 印花税率 Excel 文件（{hkex_date}）" if hkex_date \
            else "HKEx — 印花税率 Excel 文件"
        sources.append(("交易所汇率（原始数据文件）", xls_url, label))

    parts: list[str] = []
    parts.append("<h3 style='color: #2c3e50;'>数据来源</h3>")
    parts.append(
        "<table style='border-collapse: collapse; width: 100%; margin: 10px 0;'>"
        "<tr style='background: #e8f4f8;'>"
        "<th style='border: 1px solid #ddd; padding: 8px; text-align: left;'>数据项</th>"
        "<th style='border: 1px solid #ddd; padding: 8px; text-align: left;'>来源页面 URL</th>"
        "</tr>"
    )
    for item, url, label in sources:
        parts.append(
            f"<tr>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{item}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px; "
            f"word-break: break-all;'>{_link(url, label)}<br>"
            f"<span style='font-size: 12px; color: #888;'>{url}</span></td>"
            f"</tr>"
        )
    parts.append("</table>")
    return "".join(parts)


def _build_email_html(result: dict) -> str:
    """构建邮件 HTML 正文，包含公式说明和计算结果。

    Args:
        result: process_rate_export 返回的结果字典

    Returns:
        HTML 格式的邮件正文
    """
    date_str = result["date_str"]
    date_val = result["date_val"]
    custom_rows = result["custom_rows"]
    exchange_rows = result["exchange_rows"]
    bochk_hkd_raw = result["bochk_hkd_raw"]
    bochk_usd_raw = result["bochk_usd_raw"]
    hkex_data = result["hkex_data"]
    errors = result["errors"]

    date_display = date_val.strftime("%Y-%m-%d")

    parts: list[str] = []
    parts.append("<html><body style='font-family: Arial, sans-serif; font-size: 14px; color: #333;'>")

    # 标题
    parts.append(
        f"<h2 style='color: #1a5276; border-bottom: 2px solid #1a5276; "
        f"padding-bottom: 8px;'>汇率报告 — {date_display}</h2>"
    )

    # 错误提示
    if errors:
        parts.append(
            "<div style='background: #fdf2f2; border: 1px solid #e74c3c; "
            "padding: 10px; margin: 10px 0; border-radius: 4px;'>"
            "<b style='color: #e74c3c;'>⚠ 数据获取异常：</b><ul>"
        )
        for e in errors:
            parts.append(f"<li>{e}</li>")
        parts.append("</ul></div>")

    # 自定义汇率计算结果（表头含计算公式，并附 Mark down / Mark up 因子列）
    parts.append("<h3 style='color: #2c3e50;'>自定义汇率计算结果</h3>")
    parts.append(
        "<p style='color: #666; font-size: 13px;'>"
        "计算公式：<b>L (Buy) = F × Mark down</b>，"
        "<b>N (Sell) = H × Mark up</b>；"
        "F = BOCHK 客户卖出价(Bid)，H = BOCHK 客户买入价(Ask)。"
        " 来源: "
        f"{_link(BOCHK_HKDRATES_PAGE, 'BOCHK 港币电汇牌价')} / "
        f"{_link(BOCHK_USDRATES_PAGE, 'BOCHK 美元电汇牌价')}"
        "</p>"
    )
    _th = "border: 1px solid #ddd; padding: 8px; text-align: center;"
    _sub = "font-weight: normal; font-size: 11px; color: #555;"
    parts.append(
        "<table style='border-collapse: collapse; width: 100%; margin: 10px 0;'>"
        "<tr style='background: #d5f5e3;'>"
        f"<th style='{_th}'>源币种</th>"
        f"<th style='{_th}'>目标币种</th>"
        f"<th style='{_th}'>F (Bid)<br>"
        f"<span style='{_sub}'>BOCHK 客户卖出价</span></th>"
        f"<th style='{_th}'>H (Ask)<br>"
        f"<span style='{_sub}'>BOCHK 客户买入价</span></th>"
        f"<th style='{_th}'>Mark down<br>"
        f"<span style='{_sub}'>因子</span></th>"
        f"<th style='{_th}'>Mark up<br>"
        f"<span style='{_sub}'>因子</span></th>"
        f"<th style='{_th}'>L (Buy)<br>"
        f"<span style='{_sub}'>= F × Mark down</span></th>"
        f"<th style='{_th}'>N (Sell)<br>"
        f"<span style='{_sub}'>= H × Mark up</span></th>"
        f"<th style='{_th}'>计算过程</th>"
        "</tr>"
    )
    for row in custom_rows:
        pair_key = f"{row['from_ccy']}/{row['to_ccy']}"
        raw = _find_raw_rate(pair_key, bochk_hkd_raw, bochk_usd_raw)
        f_val = raw.get("bid", 0.0) if raw else 0.0
        h_val = raw.get("ask", 0.0) if raw else 0.0
        md, mu = MARKDOWN_FACTORS.get(pair_key, DEFAULT_MARKDOWN_FACTOR)
        calc = f"L = {f_val:.6f} × {md} = {row['buy']:.6f}<br>N = {h_val:.6f} × {mu} = {row['sell']:.6f}"
        parts.append(
            f"<tr>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['from_ccy']}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['to_ccy']}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{f_val:.6f}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{h_val:.6f}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{md}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px;'>{mu}</td>"
            f"<td style='border: 1px solid #ddd; padding: 8px; color: #27ae60;'><b>{row['buy']:.6f}</b></td>"
            f"<td style='border: 1px solid #ddd; padding: 8px; color: #e74c3c;'><b>{row['sell']:.6f}</b></td>"
            f"<td style='border: 1px solid #ddd; padding: 8px; font-size: 12px;'>{calc}</td>"
            f"</tr>"
        )
    parts.append("</table>")

    # 交易所汇率
    parts.append("<h3 style='color: #2c3e50;'>交易所汇率（HKEx 印花税率）</h3>")
    hkex_links = _link(HKEX_STAMPFX_URL, "HKEx 用於計算印花稅的匯率")
    hkex_xls = (hkex_data or {}).get("xls_url", "")
    if hkex_xls:
        hkex_links += f" ｜ {_link(hkex_xls, 'Excel 原始文件')}"
    parts.append(
        "<p style='color: #666; font-size: 13px;'>"
        f"来源: {hkex_links}"
        "</p>"
    )
    if exchange_rows:
        parts.append(
            "<table style='border-collapse: collapse; width: 100%; margin: 10px 0;'>"
            "<tr style='background: #fadbd8;'>"
            "<th style='border: 1px solid #ddd; padding: 8px;'>源币种</th>"
            "<th style='border: 1px solid #ddd; padding: 8px;'>目标币种</th>"
            "<th style='border: 1px solid #ddd; padding: 8px;'>单位</th>"
            "<th style='border: 1px solid #ddd; padding: 8px;'>Buy</th>"
            "<th style='border: 1px solid #ddd; padding: 8px;'>Sell</th>"
            "</tr>"
        )
        for row in exchange_rows:
            parts.append(
                f"<tr>"
                f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['from_ccy']}</td>"
                f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['to_ccy']}</td>"
                f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['unit']}</td>"
                f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['buy']:.6f}</td>"
                f"<td style='border: 1px solid #ddd; padding: 8px;'>{row['sell']:.6f}</td>"
                f"</tr>"
            )
        parts.append("</table>")
    else:
        parts.append("<p style='color: #999;'>无交易所汇率数据</p>")

    # 数据来源
    parts.append(_build_sources_html(result))

    # 附件说明
    parts.append("<h3 style='color: #2c3e50;'>附件</h3>")
    parts.append("<ul>")
    parts.append(f"<li>自定义汇率{date_str[2:]}.xlsx — BOCHK 来源，按公式计算</li>")
    if result.get("exchange_path"):
        parts.append(f"<li>交易所汇率{date_str[2:]}.xlsx — HKEx 印花税率</li>")
    parts.append("</ul>")

    parts.append(
        f"<hr style='border: none; border-top: 1px solid #ddd; margin: 20px 0;'>"
        f"<p style='color: #999; font-size: 12px;'>"
        f"此邮件由系统自动发送 — {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        f"</p>"
    )
    parts.append("</body></html>")

    return "".join(parts)


def _date_display(date_str: str) -> str:
    """将 YYYYMMDD 转为 YYYY-MM-DD 显示格式。"""
    return dt.datetime.strptime(date_str, "%Y%m%d").strftime("%Y-%m-%d")


def _build_error_email_html(
    errors: list[str],
    date_str: str,
    traceback_text: str = "",
) -> str:
    """构建错误报告邮件 HTML 正文。

    Args:
        errors: 错误信息列表
        date_str: 目标日期 YYYYMMDD
        traceback_text: 异常的 traceback 文本，可为空

    Returns:
        HTML 格式的邮件正文
    """
    parts: list[str] = []
    parts.append("<html><body style='font-family: Arial, sans-serif; font-size: 14px; color: #333;'>")
    parts.append(
        f"<h2 style='color: #c0392b; border-bottom: 2px solid #c0392b; "
        f"padding-bottom: 8px;'>⚠ 汇率报告错误报告 — {_date_display(date_str)}</h2>"
    )
    parts.append(
        "<div style='background: #fdf2f2; border: 1px solid #e74c3c; "
        "padding: 10px; margin: 10px 0; border-radius: 4px;'>"
        "<b style='color: #e74c3c;'>本次汇率获取/导出出现以下问题，请人工核查：</b>"
        "<ul>"
    )
    for e in errors:
        parts.append(f"<li>{e}</li>")
    parts.append("</ul></div>")
    parts.append(
        "<p style='color: #c0392b;'><b>本次未发送汇率邮件，也未附带任何汇率附件，"
        "请修复后重新执行。</b></p>"
    )

    parts.append("<h3 style='color: #2c3e50;'>数据来源</h3><ul>")
    parts.append(f"<li>{_link(BOCHK_HKDRATES_PAGE, 'BOCHK 港币电汇牌价')}</li>")
    parts.append(f"<li>{_link(BOCHK_USDRATES_PAGE, 'BOCHK 美元电汇牌价')}</li>")
    parts.append(f"<li>{_link(HKEX_STAMPFX_URL, 'HKEx 印花税率')}</li>")
    parts.append("</ul>")

    if traceback_text:
        parts.append("<h3 style='color: #2c3e50;'>异常堆栈</h3>")
        parts.append(
            "<pre style='background: #f7f7f7; border: 1px solid #ddd; "
            "padding: 10px; font-size: 12px; white-space: pre-wrap;'>"
            f"{traceback_text}</pre>"
        )

    parts.append(
        f"<hr style='border: none; border-top: 1px solid #ddd; margin: 20px 0;'>"
        f"<p style='color: #999; font-size: 12px;'>"
        f"此邮件由系统自动发送 — {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        f"</p>"
    )
    parts.append("</body></html>")
    return "".join(parts)


def _find_raw_rate(
    pair_key: str,
    bochk_hkd: dict[str, dict[str, float]],
    bochk_usd: dict[str, dict[str, float]],
) -> dict[str, float] | None:
    """从 BOCHK 原始汇率字典中查找指定币种对的汇率。

    Args:
        pair_key: 币种对字符串，如 "CNY/HKD" 或 "USD/CNY"
        bochk_hkd: 港币牌价字典 {币种代码: {"bid": F, "ask": H}}
        bochk_usd: 美元牌价字典 {"FROM/TO": {"bid": F, "ask": H}}

    Returns:
        {"bid": float, "ask": float} 或 None
    """
    from_ccy, to_ccy = pair_key.split("/", 1)
    if to_ccy == "HKD" and from_ccy in bochk_hkd:
        return bochk_hkd[from_ccy]
    if pair_key in bochk_usd:
        return bochk_usd[pair_key]
    return None


# -----------------------------------------------------------------------
# 主入口
# -----------------------------------------------------------------------

def main() -> None:
    """fxmail 命令入口：获取汇率 → 生成 Excel → 发送邮件。"""
    parser = argparse.ArgumentParser(description="发送当日汇率邮件")
    parser.add_argument(
        "--date", "-d",
        help="指定日期 YYYYMMDD，默认今天",
    )
    parser.add_argument(
        "--no-wechat",
        action="store_true",
        help="不发送企业微信错误通知",
    )
    args = parser.parse_args()

    date_str = args.date
    if date_str is None:
        date_str = dt.date.today().strftime("%Y%m%d")
    try:
        _date_display(date_str)
    except ValueError:
        print(f"错误: 日期格式不正确 {date_str} (应为 YYYYMMDD)", file=sys.stderr)
        sys.exit(1)

    receiver = _env("FX_RECIEVER")
    if not receiver:
        print("错误: FX_RECIEVER 未配置", file=sys.stderr)
        sys.exit(1)

    logger.info("===== FX Mail 开始, date=%s =====", date_str)

    # 1. 获取汇率并生成 Excel（异常不中断，统一收集为错误）
    result: dict | None = None
    errors: list[str] = []
    tb_text = ""
    try:
        result = process_rate_export(
            date_str=date_str,
            send_error_notify=not args.no_wechat,
        )
        errors = list(result.get("errors") or [])
    except Exception as exc:
        tb_text = traceback.format_exc()
        errors = [f"汇率处理流程执行失败: {type(exc).__name__}: {exc}"]
        logger.error("[FX] 汇率处理失败:\n%s", tb_text)
        if not args.no_wechat:
            try:
                send_error(
                    f"【汇率导出异常 {date_str}】"
                    f"{type(exc).__name__}: {exc}"
                )
            except Exception as we:
                logger.error("[FX] 企业微信错误通知发送失败: %s", we)

    # 2. 有错误则只发错误报告邮件（不发汇率邮件，也不带任何汇率附件）
    if errors:
        subject = f"【错误报告】汇率报告 — {_date_display(date_str)} 数据异常"
        html_body = _build_error_email_html(errors, date_str, tb_text)

        logger.error("发送错误报告邮件到 %s: %s", receiver, subject)
        send_mail(
            subject=subject,
            body_html=html_body,
            recipients=[receiver],
            attachments=None,
        )
        print(f"⚠ 错误报告邮件已发送到 {receiver}: {subject}", file=sys.stderr)
        logger.info("===== FX Mail 异常结束 =====")
        sys.exit(1)

    # 3. 正常：构建邮件正文并发送
    html_body = _build_email_html(result)
    subject = f"汇率报告 — {result['date_val'].strftime('%Y-%m-%d')}"
    attachments = [
        p for p in (result["custom_path"], result["exchange_path"]) if p
    ]

    logger.info("发送邮件到 %s: %s", receiver, subject)
    send_mail(
        subject=subject,
        body_html=html_body,
        recipients=[receiver],
        attachments=attachments,
    )
    print(f"✅ 邮件已发送到 {receiver}: {subject}")

    logger.info("===== FX Mail 完成 =====")


if __name__ == "__main__":
    main()
