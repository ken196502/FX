"""
ms_graph.py - Microsoft Graph API 认证与邮件读取
=================================================

使用 client credentials flow 获取 access token，
通过 Graph API 读取指定邮箱收件箱中的邮件。

所需环境变量：
  MS_TENANT_ID     - Azure AD 租户 ID
  MS_CLIENT_ID     - 应用注册的 client ID
  MS_CLIENT_SECRET - 应用注册的 client secret
  MS_MAILBOX       - 要读取的邮箱地址（如 shared mailbox）
"""

from __future__ import annotations

import datetime as dt
import logging
import os
from dataclasses import dataclass
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------

def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


MS_TENANT_ID = _env("MS_TENANT_ID")
MS_CLIENT_ID = _env("MS_CLIENT_ID")
MS_CLIENT_SECRET = _env("MS_CLIENT_SECRET")
MS_MAILBOX = _env("MS_MAILBOX") or _env("SENDER")

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_AUTH_BASE = "https://login.microsoftonline.com"


# ---------------------------------------------------------------------------
# Token
# ---------------------------------------------------------------------------

_token_cache: dict[str, str | float] = {"token": "", "expires": 0.0}


def get_access_token() -> str:
    """通过 client credentials flow 获取 Graph API access token。

    缓存 token 直到过期前 60 秒，避免频繁请求。

    需要 Mail.Read 应用级权限（application permission）。
    """
    cached = _token_cache["token"]
    expires = _token_cache["expires"]
    if cached and dt.datetime.now().timestamp() < expires:
        return cached

    if not MS_TENANT_ID or not MS_CLIENT_ID or not MS_CLIENT_SECRET:
        raise RuntimeError(
            "MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET 未配置，"
            "无法获取 access token"
        )

    url = f"{_AUTH_BASE}/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    resp = requests.post(url, data=data, timeout=30)
    if not resp.ok:
        logger.error("[MS Graph] token 请求失败: %s %s",
                      resp.status_code, resp.text)
        resp.raise_for_status()
    body = resp.json()

    token = body["access_token"]
    expires_in = int(body.get("expires_in", 3600))
    _token_cache["token"] = token
    _token_cache["expires"] = (
        dt.datetime.now().timestamp() + expires_in - 60
    )
    logger.info("[MS Graph] 获取 access token 成功，有效期 %ds", expires_in)
    return token


# ---------------------------------------------------------------------------
# 邮件数据结构
# ---------------------------------------------------------------------------

@dataclass
class MailItem:
    """一封邮件的摘要信息。"""
    message_id: str
    subject: str
    sender: str
    sent_datetime: str
    received_datetime: str
    has_attachments: bool


# ---------------------------------------------------------------------------
# 邮件查询
# ---------------------------------------------------------------------------

HK_TZ = dt.timezone(dt.timedelta(hours=8))


def _build_date_filter(date_str: str) -> str:
    """构造 Graph API $filter 的日期范围条件。

    按香港时间（UTC+8）当天 00:00:00 ~ 23:59:59 转换为 UTC 时间查询。
    即 UTC 前一天 16:00:00 ~ 当天 15:59:59。

    Args:
        date_str: YYYYMMDD 格式日期
    """
    d = dt.date.fromisoformat(date_str)
    hk_start = dt.datetime(d.year, d.month, d.day, tzinfo=HK_TZ)
    hk_end = dt.datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=HK_TZ)
    utc_start = hk_start.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    utc_end = hk_end.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return f"sentDateTime ge {utc_start} and sentDateTime le {utc_end}"


def fetch_mails(
    sender: str | None = None,
    date_str: str | None = None,
    mailbox: str | None = None,
    top: int = 50,
) -> list[MailItem]:
    """从指定邮箱收件箱获取邮件列表。

    优先使用服务端发件人过滤 (from/emailAddress/address eq '...') 缩小结果集，
    避免客户端过滤和分页问题。

    Args:
        sender: 发件人邮箱地址过滤（精确匹配，不区分大小写）。None 表示不过滤。
        date_str: 发送日期过滤 (YYYYMMDD)。None 表示不过滤。
        mailbox: 要读取的邮箱地址。None 则使用 MS_MAILBOX 环境变量。
        top: 最多返回的邮件数。

    Returns:
        MailItem 列表，按发送时间降序排列。
    """
    token = get_access_token()
    mbox = mailbox or MS_MAILBOX
    if not mbox:
        raise RuntimeError("未指定邮箱地址 (MS_MAILBOX 未配置且未传参)")

    headers = {"Authorization": f"Bearer {token}"}
    url = f"{_GRAPH_BASE}/users/{mbox}/mailFolders/inbox/messages"

    params: dict[str, str] = {
        "$select": (
            "id,subject,from,sentDateTime,receivedDateTime,hasAttachments"
        ),
        "$orderby": "sentDateTime desc",
        "$top": str(top),
    }
    filters: list[str] = []
    if date_str:
        filters.append(_build_date_filter(date_str))
    if sender:
        filters.append(f"from/emailAddress/address eq '{sender}'")
    if filters:
        params["$filter"] = " and ".join(filters)

    logger.info(
        "[MS Graph] 查询邮箱=%s, 发件人=%s, 日期=%s",
        mbox, sender or "(全部)", date_str or "(全部)",
    )
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    resp.raise_for_status()
    body = resp.json()

    items: list[MailItem] = []
    for msg in body.get("value", []):
        from_obj = msg.get("from", {})
        from_email_obj = from_obj.get("emailAddress", {})
        sender_addr = from_email_obj.get("address", "")
        items.append(MailItem(
            message_id=msg.get("id", ""),
            subject=msg.get("subject", "(无主题)"),
            sender=sender_addr,
            sent_datetime=msg.get("sentDateTime", ""),
            received_datetime=msg.get("receivedDateTime", ""),
            has_attachments=msg.get("hasAttachments", False),
        ))

    logger.info("[MS Graph] 查询到 %d 封邮件", len(items))
    return items


def fetch_mail_body(message_id: str, mailbox: str | None = None) -> str:
    """获取单封邮件的纯文本正文。

    Args:
        message_id: 邮件 ID
        mailbox: 邮箱地址，None 则使用 MS_MAILBOX
    """
    token = get_access_token()
    mbox = mailbox or MS_MAILBOX
    if not mbox:
        raise RuntimeError("未指定邮箱地址 (MS_MAILBOX 未配置且未传参)")

    headers = {"Authorization": f"Bearer {token}"}
    url = f"{_GRAPH_BASE}/users/{mbox}/messages/{message_id}"
    params = {"$select": "subject,body"}

    resp = requests.get(url, headers=headers, params=params, timeout=30)
    resp.raise_for_status()
    body = resp.json()
    content = body.get("body", {})
    content_type = content.get("contentType", "")
    content_str = content.get("content", "")

    if content_type.lower() == "html":
        content_str = _strip_html(content_str)

    return content_str.strip()


def fetch_mail_attachments(
    message_id: str,
    mailbox: str | None = None,
) -> list[dict]:
    """获取邮件附件列表（含元数据，不含内容）。

    Args:
        message_id: 邮件 ID
        mailbox: 邮箱地址，None 则使用 MS_MAILBOX

    Returns:
        附件字典列表，每个含 id, name, contentType, size, contentBytes 等
    """
    token = get_access_token()
    mbox = mailbox or MS_MAILBOX
    if not mbox:
        raise RuntimeError("未指定邮箱地址 (MS_MAILBOX 未配置且未传参)")

    headers = {"Authorization": f"Bearer {token}"}
    url = f"{_GRAPH_BASE}/users/{mbox}/messages/{message_id}/attachments"

    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    body = resp.json()
    return body.get("value", [])


def fetch_attachment_content(
    message_id: str,
    attachment_id: str,
    mailbox: str | None = None,
) -> bytes:
    """下载邮件附件的原始二进制内容。

    Args:
        message_id: 邮件 ID
        attachment_id: 附件 ID
        mailbox: 邮箱地址，None 则使用 MS_MAILBOX

    Returns:
        附件的原始字节内容
    """
    token = get_access_token()
    mbox = mailbox or MS_MAILBOX
    if not mbox:
        raise RuntimeError("未指定邮箱地址 (MS_MAILBOX 未配置且未传参)")

    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{_GRAPH_BASE}/users/{mbox}/messages/{message_id}"
        f"/attachments/{attachment_id}/$value"
    )

    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    return resp.content


def _strip_html(html: str) -> str:
    """简易 HTML 标签移除，提取纯文本。"""
    import re
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"</?p[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&")
    text = text.replace("&lt;", "<").replace("&gt;", ">")
    import html as html_mod
    text = html_mod.unescape(text)
    lines = [line.strip() for line in text.split("\n")]
    text = "\n".join(line for line in lines if line)
    return text
