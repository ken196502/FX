"""
ms_mail.py - 通过 Microsoft Graph API 发送邮件
================================================

使用 client credentials flow 获取 access token，
通过 Graph API 以指定邮箱身份发送邮件（含附件）。

所需环境变量（通过 ms_graph 模块读取）：
  MS_TENANT_ID     - Azure AD 租户 ID
  MS_CLIENT_ID     - 应用注册的 client ID
  MS_CLIENT_SECRET - 应用注册的 client secret
  SENDER           - 发件邮箱地址
  FX_RECIEVER      - 收件邮箱地址（完整汇率报告），也支持 FX_RECEIVER 拼写
  BOC_RECIEVER     - 收件邮箱地址（仅自定义汇率，BOCHK 来源），
                     也支持 BOC_RECEIVER 拼写

收件人环境变量支持多个邮箱地址，用 , 或 ; 分隔，
统一转换为 Graph API 的 toRecipients 列表（即一封邮件多个收件人）。
"""

from __future__ import annotations

import base64
from pathlib import Path

from log_setup import get_logger
from ms_graph import get_access_token, MS_MAILBOX

logger = get_logger(__name__)

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _read_file_b64(path: Path) -> str:
    """读取文件并返回 base64 编码字符串。

    Args:
        path: 文件路径

    Returns:
        base64 编码的文件内容字符串
    """
    data = path.read_bytes()
    return base64.b64encode(data).decode("ascii")


def _build_attachment(path: Path) -> dict:
    """构建 Graph API 邮件附件 JSON 结构。

    Args:
        path: 附件文件路径

    Returns:
        附件字典，含 @odata.type / name / contentType / contentBytes
    """
    if not path.exists():
        logger.error("[MS Mail] 附件文件不存在: %s", path)
    else:
        logger.debug(
            "[MS Mail] 读取附件: %s (%.0f KB)",
            path.name, path.stat().st_size / 1024,
        )
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": path.name,
        "contentType": _guess_content_type(path.suffix),
        "contentBytes": _read_file_b64(path),
    }


def _guess_content_type(suffix: str) -> str:
    """根据文件扩展名猜测 MIME 类型。

    Args:
        suffix: 文件扩展名（含点，如 .xlsx）

    Returns:
        MIME 类型字符串
    """
    mapping = {
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xls": "application/vnd.ms-excel",
        ".csv": "text/csv",
        ".pdf": "application/pdf",
        ".txt": "text/plain",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
    }
    return mapping.get(suffix.lower(), "application/octet-stream")


def send_mail(
    subject: str,
    body_html: str,
    recipients: list[str],
    attachments: list[Path] | None = None,
    sender: str | None = None,
) -> dict:
    """通过 Microsoft Graph API 发送邮件。

    使用 client credentials flow，以 sender 邮箱身份发送。
    需要 Mail.Send 应用级权限（application permission）。

    Args:
        subject: 邮件主题
        body_html: 邮件正文 HTML 内容
        recipients: 收件人邮箱地址列表
        attachments: 附件文件路径列表
        sender: 发件邮箱地址，None 则使用 MS_MAILBOX/SENDER

    Returns:
        Graph API 返回的 JSON 字典
    """
    token = get_access_token()
    mbox = sender or MS_MAILBOX
    if not mbox:
        raise RuntimeError("未指定发件邮箱地址 (MS_MAILBOX/SENDER 未配置)")

    recipients = [a.strip() for a in (recipients or []) if a and a.strip()]
    if not recipients:
        raise ValueError("收件人为空，无法发送邮件")

    to_recipients = [
        {"emailAddress": {"address": addr}}
        for addr in recipients
    ]

    message: dict = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": body_html,
        },
        "toRecipients": to_recipients,
    }

    if attachments:
        message["attachments"] = [_build_attachment(p) for p in attachments]

    url = f"{_GRAPH_BASE}/users/{mbox}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    payload = {"message": message, "saveToSentItems": True}

    logger.info(
        "[MS Mail] 发送邮件: from=%s, to=%s, subject=%s, attachments=%d",
        mbox, recipients, subject, len(attachments or []),
    )
    logger.debug("[MS Mail] 请求地址: %s", url)
    for path in attachments or []:
        path = Path(path)
        if not path.exists():
            logger.warning("[MS Mail] 附件不存在，将被跳过: %s", path)
        else:
            logger.debug(
                "[MS Mail] 附件: %s (%.0f KB)",
                path.name, path.stat().st_size / 1024,
            )

    import requests
    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=60)
    except Exception as exc:
        logger.error(
            "[MS Mail] 邮件发送请求异常 (%s: %s)", type(exc).__name__, exc
        )
        raise
    if not resp.ok:
        logger.error(
            "[MS Mail] 邮件发送失败: %s %s", resp.status_code, resp.text[:500]
        )
        resp.raise_for_status()

    logger.info("[MS Mail] 邮件发送成功 (HTTP %s)", resp.status_code)
    return {"status": "ok", "status_code": resp.status_code}
