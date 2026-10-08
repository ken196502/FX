"""
log_setup.py - 统一日志配置
=============================

项目此前只有 `main.py` 里的 logging.basicConfig（仅输出到控制台），
运行过程没有任何落盘记录，出问题后无法回溯。本模块提供统一入口：

    from log_setup import get_logger, setup_logging

    logger = get_logger(__name__)

默认行为：
  - 控制台：输出到 stdout，便于命令行/任务计划里直接查看
  - 文件：写到 <项目根>/logs/fx_YYYYMMDD.log，UTF-8，
    单文件 5MB 后轮转，最多保留 10 个历史文件

环境变量（可选）：
  LOG_LEVEL - 日志级别，默认 INFO，支持 DEBUG/INFO/WARNING/ERROR/CRITICAL
  LOG_DIR   - 日志目录，默认 <项目根>/logs
  LOG_FILE_PREFIX - 日志文件名前缀，默认 fx（最终 fx_YYYYMMDD.log）
"""

from __future__ import annotations

import logging
import os
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

from dotenv import load_dotenv

_PROJECT_ROOT = Path(__file__).resolve().parent
load_dotenv(_PROJECT_ROOT / ".env")

DEFAULT_LOG_DIR = _PROJECT_ROOT / "logs"
DEFAULT_FILE_PREFIX = "fx"

_MAX_BYTES = 5 * 1024 * 1024  # 单个日志文件 5MB
_BACKUP_COUNT = 10  # 最多保留 10 个历史文件

_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# 第三方库日志过于嘈杂，默认压到 WARNING
_NOISY_LOGGERS = (
    "urllib3",
    "requests",
    "httpx",
    "httpcore",
    "msal",
    "azure",
    "asyncio",
    "playwright",
    "openpyxl",
)

_configured = False
_log_file: Path | None = None

# 控制台在 Windows 下可能是 GBK，中文写不进去时降级处理，避免抛异常
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:  # pragma: no cover - 极端环境忽略
        pass


def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


def _resolve_level(value, default: int = logging.INFO) -> int:
    """将字符串/整数转换为 logging 级别，无法识别时返回默认值。

    Args:
        value: 级别字符串（如 "DEBUG"）或 logging 级别整数
        default: 无法识别时的默认级别

    Returns:
        logging 级别整数
    """
    if value is None or value == "":
        return default
    if isinstance(value, int):
        return value
    level = logging.getLevelName(str(value).strip().upper())
    return level if isinstance(level, int) else default


def setup_logging(
    level: str | int | None = None,
    log_dir: str | Path | None = None,
    console: bool = True,
    force: bool = False,
) -> Path | None:
    """初始化根 logger：控制台 + 按大小轮转的文件日志。

    幂等：重复调用不会重复添加 handler。需要按命令行参数重新配置时
    传 force=True（main.py 解析完 --log-level/--log-dir 后使用）。

    Args:
        level: 日志级别，None 时读 LOG_LEVEL 环境变量，默认 INFO
        log_dir: 日志目录，None 时读 LOG_DIR 环境变量，默认 ./logs
        console: 是否输出到控制台
        force: True 时先移除已有 handler 再重新配置

    Returns:
        当前日志文件路径；文件日志不可用时返回 None
    """
    global _configured, _log_file

    root = logging.getLogger()
    if _configured and not force:
        return _log_file

    if force:
        for handler in list(root.handlers):
            root.removeHandler(handler)

    resolved_level = _resolve_level(level or _env("LOG_LEVEL"))
    prefix = _env("LOG_FILE_PREFIX") or DEFAULT_FILE_PREFIX
    directory = Path(log_dir or _env("LOG_DIR") or DEFAULT_LOG_DIR)

    formatter = logging.Formatter(_LOG_FORMAT, datefmt=_DATE_FORMAT)
    root.setLevel(resolved_level)

    if console:
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setFormatter(formatter)
        stream_handler.setLevel(resolved_level)
        root.addHandler(stream_handler)

    file_path = directory / f"{prefix}_{datetime.now().strftime('%Y%m%d')}.log"
    try:
        directory.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            file_path,
            maxBytes=_MAX_BYTES,
            backupCount=_BACKUP_COUNT,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        file_handler.setLevel(resolved_level)
        root.addHandler(file_handler)
        _log_file = file_path
    except Exception as exc:
        # 目录不可写时不能让整个程序挂掉，只用控制台继续
        _log_file = None
        root.warning("日志文件初始化失败，仅输出到控制台: %s (%s)", file_path, exc)

    if resolved_level > logging.DEBUG:
        for name in _NOISY_LOGGERS:
            logging.getLogger(name).setLevel(logging.WARNING)

    _configured = True
    logging.getLogger(__name__).info(
        "日志已初始化: 级别=%s, 文件=%s",
        logging.getLevelName(resolved_level),
        _log_file or "(仅控制台)",
    )
    return _log_file


def get_logger(name: str | None = None) -> logging.Logger:
    """获取已配置好的 logger（首次调用会自动初始化日志）。

    Args:
        name: logger 名称，通常传 __name__

    Returns:
        配置完成的 logging.Logger
    """
    setup_logging()
    return logging.getLogger(name or "fx")


def log_file_path() -> Path | None:
    """返回当前日志文件路径，未启用文件日志时为 None。"""
    return _log_file


if __name__ == "__main__":
    setup_logging(force=True)
    get_logger(__name__).info("log_setup 自测日志输出正常")
    print(f"日志文件: {log_file_path()}")
