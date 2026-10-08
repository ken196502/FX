"""
sc_infra.py - 企业微信 Webhook 通知
====================================

通过企业微信机器人 Webhook 发送文本消息和文件。

所需环境变量：
  ERROR_REPORT_WEBHOOK_KEY - Webhook 的 key 部分
"""

from __future__ import annotations

import os
from pathlib import Path

import requests
from dotenv import load_dotenv

from log_setup import get_logger

load_dotenv(Path(__file__).resolve().parent / ".env")

logger = get_logger(__name__)

_WECHAT_WEBHOOK_BASE = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send"
_WECHAT_UPLOAD_BASE = "https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media"


def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


def _webhook_key() -> str:
    key = _env("ERROR_REPORT_WEBHOOK_KEY")
    if not key:
        logger.error("[WeChat] ERROR_REPORT_WEBHOOK_KEY 未配置，无法发送通知")
        raise RuntimeError("ERROR_REPORT_WEBHOOK_KEY 未配置")
    return key


def _post_wechat(
    url: str,
    params: dict,
    label: str,
    json_payload: dict | None = None,
    files=None,
    timeout: int = 30,
) -> dict:
    """发送企业微信 Webhook 请求并统一记录日志。

    Args:
        url: Webhook 地址
        params: URL 参数（含 key，日志中会打码）
        label: 日志标签，如「文本消息」
        json_payload: JSON 请求体
        files: 文件上传用的 files 参数
        timeout: 超时秒数

    Returns:
        API 返回的 JSON 字典
    """
    safe_params = {**params}
    if "key" in safe_params:
        safe_params["key"] = f"{str(safe_params['key'])[:6]}***"
    logger.debug("[WeChat] 请求 %s: %s params=%s", label, url, safe_params)
    try:
        resp = requests.post(
            url, params=params, json=json_payload, files=files, timeout=timeout,
        )
    except Exception as exc:
        logger.error(
            "[WeChat] %s 请求异常 (%s: %s)", label, type(exc).__name__, exc
        )
        raise
    if not resp.ok:
        logger.error(
            "[WeChat] %s 请求失败: %s %s", label, resp.status_code,
            resp.text[:300],
        )
        resp.raise_for_status()
    result = resp.json()
    logger.info("[WeChat] %s 发送结果: %s", label, result)
    return result


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
    logger.info("[WeChat] 发送文本消息 (%d 字符)", len(text))
    logger.debug("[WeChat] 文本消息内容: %s", text)
    return _post_wechat(
        _WECHAT_WEBHOOK_BASE,
        params={"key": key},
        label="文本消息",
        json_payload=payload,
        timeout=30,
    )


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

    logger.info(
        "[WeChat] 上传文件: %s (%.0f KB)",
        file_path.name, file_path.stat().st_size / 1024,
    )

    # 上传文件获取 media_id
    with open(file_path, "rb") as f:
        upload_data = _post_wechat(
            _WECHAT_UPLOAD_BASE,
            params={"key": key, "type": "file"},
            label=f"文件上传({file_path.name})",
            files={"media": (file_path.name, f)},
            timeout=60,
        )
    media_id = upload_data.get("media_id")
    if not media_id:
        logger.error("[WeChat] 文件上传未返回 media_id: %s", upload_data)
        raise RuntimeError(f"上传文件失败: {upload_data}")
    logger.info("[WeChat] 文件上传成功: %s -> media_id=%s", file_path.name, media_id)

    # 发送文件消息
    payload = {
        "msgtype": "file",
        "file": {"media_id": media_id},
    }
    return _post_wechat(
        _WECHAT_WEBHOOK_BASE,
        params={"key": key},
        label=f"文件消息({file_path.name})",
        json_payload=payload,
        timeout=30,
    )


def send_error(text: str) -> dict:
    """发送错误通知到企业微信 Webhook。

    Args:
        text: 错误信息文本

    Returns:
        API 返回的 JSON 字典
    """
    logger.warning("[WeChat] 发送错误通知: %s", text)
    return send_wechat(f"❌ {text}")
