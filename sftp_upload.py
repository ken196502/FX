"""
sftp_upload.py - 通过 SFTP 上传汇率 Excel
=========================================

读取 .env 中的 SFTP 配置，把生成的 Excel 上传到 SFTP 目录下的
日期子目录（YYYYMMDD）：

    <SFTP_DIR>/<YYYYMMDD>/<文件名>

例如 SFTP_DIR=FX_Rates、日期 20261009 时，自定义汇率261009.xlsx
上传到：

    FX_Rates/20261009/自定义汇率261009.xlsx

所需环境变量：
  SFTP      - SFTP 地址 host[:port]，如 10.202.5.244:2022（端口默认 22）
  SFTP_DIR  - SFTP 上的根目录，如 FX_Rates（也支持 /abs/path 绝对路径）
  SFTP_USER - 登录用户名
  SFTP_PWD  - 登录密码

未配置 SFTP 相关变量时上传会直接跳过（不抛异常），避免影响主流程；
已配置但连接/上传失败时抛出 RuntimeError，由调用方决定如何处理。
"""

from __future__ import annotations

import os
import posixpath
from pathlib import Path

from dotenv import load_dotenv

from log_setup import get_logger

load_dotenv(Path(__file__).resolve().parent / ".env")

logger = get_logger(__name__)

# SFTP 默认端口（SFTP 未写端口时使用）
DEFAULT_SFTP_PORT = 22

# 连接 / 上传超时（秒）
_SFTP_TIMEOUT = 30


def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


def sftp_enabled() -> bool:
    """判断 .env 中是否配置了可用的 SFTP（地址与用户名同时存在即可）。

    Returns:
        True 表示已配置 SFTP，可以尝试上传
    """
    return bool(_env("SFTP") and _env("SFTP_USER"))


def _split_host_port(raw: str) -> tuple[str, int]:
    """解析 host[:port] 形式的 SFTP 地址。

    Args:
        raw: 如 "10.202.5.244:2022" 或 "10.202.5.244"

    Returns:
        (host, port) 元组，端口缺省为 22

    Raises:
        RuntimeError: 地址为空或端口不是数字
    """
    raw = raw.strip()
    if not raw:
        raise RuntimeError("SFTP 未配置（SFTP 为空）")
    if ":" in raw:
        host, _, port_str = raw.rpartition(":")
        try:
            port = int(port_str)
        except ValueError:
            raise RuntimeError(f"SFTP 端口不是数字: {raw}")
    else:
        host, port = raw, DEFAULT_SFTP_PORT
    return host, port


def _ensure_remote_dir(sftp, remote_dir: str) -> str:
    """确保 SFTP 上的目录存在，不存在则逐级创建。

    Args:
        sftp: paramiko 的 SFTPClient
        remote_dir: 远端目录（相对路径按 SFTP 登录后的默认目录解析，
            也支持以 / 开头的绝对路径）

    Returns:
        规范化后的远端目录路径
    """
    parts = [p for p in remote_dir.split("/") if p]
    if not parts:
        return "/"
    cur = "/" if remote_dir.startswith("/") else ""
    for part in parts:
        cur = posixpath.join(cur, part) if cur else part
        try:
            sftp.stat(cur)
        except OSError:
            sftp.mkdir(cur)
            logger.info("[SFTP] 创建远端目录: %s", cur)
    return cur


def upload_file(local_path: Path, remote_path: str) -> str:
    """把本地文件上传到 SFTP 指定路径（自动创建远端目录）。

    Args:
        local_path: 本地文件路径
        remote_path: 远端完整路径（含文件名），目录不存在时自动创建

    Returns:
        实际写入的远端路径

    Raises:
        RuntimeError: SFTP 未配置、连接失败或上传失败
        FileNotFoundError: 本地文件不存在
    """
    local_path = Path(local_path)
    if not local_path.exists():
        raise FileNotFoundError(f"本地文件不存在: {local_path}")

    if not sftp_enabled():
        raise RuntimeError("SFTP 未配置（缺少 SFTP / SFTP_USER），跳过上传")

    import paramiko  # 延迟导入：未使用 SFTP 时不必依赖

    host, port = _split_host_port(_env("SFTP"))
    user = _env("SFTP_USER")
    pwd = _env("SFTP_PWD")
    remote_path = remote_path.replace("\\", "/")

    logger.info(
        "[SFTP] 上传 %s -> %s@%s:%d:%s (%.0f KB)",
        local_path.name, user, host, port, remote_path,
        local_path.stat().st_size / 1024,
    )

    client = paramiko.SSHClient()
    client.load_system_host_keys()
    # 内网 SFTP 服务器通常不在 known_hosts 中：首次连接自动记录并告警
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        client.connect(
            hostname=host,
            port=port,
            username=user,
            password=pwd,
            look_for_keys=False,
            allow_agent=False,
            timeout=_SFTP_TIMEOUT,
        )
        sftp = client.open_sftp()
        try:
            remote_dir = posixpath.dirname(remote_path)
            if remote_dir:
                _ensure_remote_dir(sftp, remote_dir)
            sftp.put(str(local_path), remote_path)
            try:
                size = sftp.stat(remote_path).st_size
            except OSError:
                size = local_path.stat().st_size
            logger.info(
                "[SFTP] 上传完成: %s (%d 字节)", remote_path, size,
            )
        finally:
            sftp.close()
    except Exception as exc:
        logger.error(
            "[SFTP] 上传失败 (%s -> %s): %s: %s",
            local_path.name, remote_path, type(exc).__name__, exc,
        )
        raise RuntimeError(
            f"SFTP 上传失败 ({local_path.name} -> {remote_path}): "
            f"{type(exc).__name__}: {exc}"
        ) from exc
    finally:
        client.close()

    return remote_path


def upload_rate_file(
    file_path: Path,
    date_str: str,
    remote_base: str | None = None,
) -> str:
    """把汇率 Excel 上传到 SFTP 的 <根目录>/<YYYYMMDD>/ 下。

    Args:
        file_path: 本地 Excel 文件路径，远端沿用其文件名
        date_str: 目标日期 YYYYMMDD，用作 SFTP 上的子目录名
        remote_base: SFTP 根目录，None 时取环境变量 SFTP_DIR（默认空，
            表示上传到登录后的默认目录）

    Returns:
        实际写入的远端完整路径

    Raises:
        RuntimeError: SFTP 未配置、连接失败或上传失败
        FileNotFoundError: 本地文件不存在
    """
    base = (remote_base if remote_base is not None else _env("SFTP_DIR")).strip()
    base = base.replace("\\", "/").strip("/")
    remote_path = "/".join(p for p in (base, date_str, Path(file_path).name) if p)
    return upload_file(file_path, remote_path)
