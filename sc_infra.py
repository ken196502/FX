"""
sc_infra.py - 企业微信 Webhook 通知
====================================

通过企业微信机器人 Webhook 发送文本消息和文件。

所需环境变量：
  ERROR_REPORT_WEBHOOK_KEY - Webhook 的 key 部分
"""

from __future__ import annotations

import logging
import os
from pathlib import Path

import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

logger = logging.getLogger(__name__)

_WECHAT_WEBHOOK_BASE = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send"
_WECHAT_UPLOAD_BASE = "https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media"


def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


def _webhook_key() -> str:
    key = _env("ERROR_REPORT_WEBHOOK_KEY")
    if not key:
        raise RuntimeError("ERROR_REPORT_WEBHOOK_KEY 未配置")
    return key


def send_wechat(text: str) -> dict:
    """发送文本消息到企业微信 Webhook。

    Args:
        text: 要发送的文本内容

    Returns:
        API 返回的 JSON 字典
    """
    key = _webhook_key()
    payload = {
        "msgtype": "text",
        "text": {"content": text},
    }
    resp = requests.post(
        _WECHAT_WEBHOOK_BASE,
        params={"key": key},
        json=payload,
        timeout=30,
    )
    resp.raise_for_status()
    result = resp.json()
    logger.info("[WeChat] 文本消息发送结果: %s", result)
    return result


def send_wechat_file(file_path: Path) -> dict:
    """通过企业微信 Webhook 发送文件。

    先上传文件获取 media_id，再发送文件类型消息。

    Args:
        file_path: 文件路径

    Returns:
        API 返回的 JSON 字典
    """
    key = _webhook_key()
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")

    # 上传文件获取 media_id
    with open(file_path, "rb") as f:
        upload_resp = requests.post(
            _WECHAT_UPLOAD_BASE,
            params={"key": key, "type": "file"},
            files={"media": (file_path.name, f)},
            timeout=60,
        )
    upload_resp.raise_for_status()
    upload_data = upload_resp.json()
    media_id = upload_data.get("media_id")
    if not media_id:
        raise RuntimeError(f"上传文件失败: {upload_data}")
    logger.info("[WeChat] 文件上传成功: %s -> media_id=%s", file_path.name, media_id)

    # 发送文件消息
    payload = {
        "msgtype": "file",
        "file": {"media_id": media_id},
    }
    resp = requests.post(
        _WECHAT_WEBHOOK_BASE,
        params={"key": key},
        json=payload,
        timeout=30,
    )
    resp.raise_for_status()
    result = resp.json()
    logger.info("[WeChat] 文件消息发送结果: %s", result)
    return result


def send_error(text: str) -> dict:
    """发送错误通知到企业微信 Webhook。

    Args:
        text: 错误信息文本

    Returns:
        API 返回的 JSON 字典
    """
    return send_wechat(f"❌ {text}")
