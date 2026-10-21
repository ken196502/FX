r"""
smb_upload.py - 通过 SMB 上传汇率 Excel
========================================

读取 .env 中的 SMB 配置，把生成的汇率 Excel 上传到 SMB 共享目录下的
年/日期子目录（YYYY/YYYYMMDD）：

    <SMB_REMOTE_DIR>/<YYYY>/<YYYYMMDD>/<文件名>

例如 SMB_SERVER=\\10.202.5.244\FX_Rates、日期 20261009 时，自定义汇率261009.xlsx
上传到：

    FX_Rates/2026/20261009/自定义汇率261009.xlsx

所需环境变量：
  SMB_SERVER   - 服务器地址，可写成 "10.202.5.244" / "10.202.5.244/FX_Rates"
                 / "\\10.202.5.244\FX_Rates"（也兼容旧名 SMB）
  SMB_SHARE    - 共享名（SMB_SERVER 中未带共享名时必填），如 FX_Rates
  SMB_REMOTE_DIR - 共享内的根目录，如 FX_Rates/import（缺省表示共享根目录，
                   也兼容旧名 SMB_DIR）
  SMB_USERNAME - 登录用户名（也兼容旧名 SMB_USER）
  SMB_PASSWORD - 登录密码（也兼容旧名 SMB_PWD）

未配置 SMB 相关变量时上传会直接跳过（不抛异常），由调用方决定是否改为
抓取汇率；已配置但连接/上传失败时抛出 RuntimeError，不做静默兜底。

「不重复上传」：上传前先检查远端是否已存在同名文件——
  - 远端不存在：上传；
  - 远端已存在且字节数与本地一致：跳过，不重复上传；
  - 远端已存在但字节数不一致：重新上传（上一次可能上传不完整）。
"""

from __future__ import annotations

import errno
import os
import uuid
from contextlib import contextmanager
from pathlib import Path

from dotenv import load_dotenv

from log_setup import get_logger

load_dotenv(Path(__file__).resolve().parent / ".env")

logger = get_logger(__name__)

# SMB 默认端口
DEFAULT_SMB_PORT = 445

# 连接 / 上传超时（秒）
_SMB_TIMEOUT = 30


def _env(key: str) -> str:
    return os.getenv(key, "").strip().strip('"').strip("'")


def _env_first(*keys: str) -> str:
    """按优先级读取环境变量，返回第一个非空值（兼容新旧变量名）。"""
    for key in keys:
        val = _env(key)
        if val:
            return val
    return ""


def _smb_server() -> str:
    return _env_first("SMB_SERVER", "SMB")


def _smb_user() -> str:
    return _env_first("SMB_USERNAME", "SMB_USER")


def _smb_pwd() -> str:
    return _env_first("SMB_PASSWORD", "SMB_PWD")


def _smb_remote_dir() -> str:
    return _env_first("SMB_REMOTE_DIR", "SMB_DIR")


def smb_enabled() -> bool:
    """判断 .env 中是否配置了可用的 SMB（地址与用户名同时存在即可）。

    Returns:
        True 表示已配置 SMB，可以尝试上传
    """
    return bool(_smb_server() and _smb_user())


def _parse_smb_target() -> tuple[str, str]:
    """解析 SMB 服务器地址与共享名。

    SMB 支持 "host" / "host/share" / "\\\\host\\share" 三种写法；
    只写 host 时再从 SMB_SHARE 取共享名。

    Returns:
        (server, share) 元组

    Raises:
        RuntimeError: SMB 为空或无法确定共享名
    """
    raw = _smb_server().replace("\\", "/").strip().strip("/")
    if not raw:
        raise RuntimeError("SMB 未配置（SMB_SERVER 为空）")
    parts = [p for p in raw.split("/") if p]
    if len(parts) >= 2:
        return parts[0], parts[1]
    share = _env("SMB_SHARE").strip()
    if not share:
        raise RuntimeError(
            "无法确定 SMB 共享名：请把 SMB 写成 "
            "\\\\服务器\\共享名（或 服务器/共享名），或单独配置 SMB_SHARE"
        )
    return parts[0], share


@contextmanager
def _open_tree():
    """建立 SMB 连接、登录并连上共享目录，退出时自动断开。

    Yields:
        smbprotocol.tree.TreeConnect 对象

    Raises:
        RuntimeError: SMB 未配置、连接/登录/挂载共享失败
    """
    if not smb_enabled():
        raise RuntimeError("SMB 未配置（缺少 SMB / SMB_USER）")

    # 延迟导入：未使用 SMB 时不必依赖 smbprotocol
    from smbprotocol.connection import Connection
    from smbprotocol.session import Session
    from smbprotocol.tree import TreeConnect

    server, share = _parse_smb_target()
    user = _smb_user()
    pwd = _smb_pwd()

    connection = Connection(uuid.uuid4(), server, port=DEFAULT_SMB_PORT)
    try:
        connection.connect(timeout=_SMB_TIMEOUT)
        # require_encryption=False：部分内网 SMB 服务器（如未开加密的
        # 共享）协商时不提供加密能力，强制加密会直接连不上
        session = Session(
            connection, username=user, password=pwd,
            require_encryption=False,
        )
        session.connect()
        tree = TreeConnect(session, f"\\\\{server}\\{share}")
        tree.connect()
        logger.debug("[SMB] 已连接 \\\\%s\\%s (用户 %s)", server, share, user)
        try:
            yield tree
        finally:
            tree.disconnect()
            session.disconnect()
            connection.disconnect()
    except Exception as exc:
        try:
            connection.disconnect()
        except Exception:
            pass
        raise RuntimeError(
            f"SMB 连接/登录失败 (\\\\{server}\\{share}, {user}): "
            f"{type(exc).__name__}: {exc}"
        ) from exc


def _is_not_found(exc: Exception) -> bool:
    """判断异常是否表示「远端对象不存在」。

    smbprotocol 对不同操作抛不同的异常类型：create 打不开不存在对象时抛
    ObjectNameNotFound / ObjectPathNotFound / NotFound（SMBResponseException
    家族），文件读写类操作失败时抛 SMBOSError(errno=ENOENT)，都要认。

    Args:
        exc: 捕获到的异常

    Returns:
        True 表示对象不存在
    """
    from smbprotocol.exceptions import (
        NotFound,
        ObjectNameNotFound,
        ObjectPathNotFound,
        SMBOSError,
    )

    if isinstance(exc, (ObjectNameNotFound, ObjectPathNotFound, NotFound)):
        return True
    return isinstance(exc, SMBOSError) and exc.errno == errno.ENOENT


def _is_already_exists(exc: Exception) -> bool:
    """判断异常是否表示「远端对象已存在」（如并发创建目录）。"""
    from smbprotocol.exceptions import ObjectNameCollision, SMBOSError

    if isinstance(exc, ObjectNameCollision):
        return True
    return isinstance(exc, SMBOSError) and exc.errno == errno.EEXIST


def _open_file_stat(tree, remote_path: str, directory: bool):
    """以只读属性权限打开远端文件/目录，返回其字节数。

    仅用于存在性探测：打开成功即视为存在，返回 end_of_file；
    服务端返回「对象不存在」时返回 None；其它错误（权限、网络等）
    直接抛出，不做兜底。

    Args:
        tree: TreeConnect 对象
        remote_path: 远端路径（共享内相对路径）
        directory: True 按目录探测，False 按文件探测

    Returns:
        远端对象字节数，不存在时返回 None

    Raises:
        Exception: 打开失败且不是「对象不存在」时原样抛出
    """
    from smbprotocol.open import (
        CreateDisposition,
        CreateOptions,
        FileAttributes,
        FilePipePrinterAccessMask,
        ImpersonationLevel,
        Open,
        ShareAccess,
    )

    fh = Open(tree, remote_path)
    try:
        fh.create(
            impersonation_level=ImpersonationLevel.Impersonation,
            desired_access=FilePipePrinterAccessMask.FILE_READ_ATTRIBUTES,
            file_attributes=FileAttributes.FILE_ATTRIBUTE_NORMAL,
            share_access=(
                ShareAccess.FILE_SHARE_READ
                | ShareAccess.FILE_SHARE_WRITE
                | ShareAccess.FILE_SHARE_DELETE
            ),
            create_disposition=CreateDisposition.FILE_OPEN,
            create_options=(
                CreateOptions.FILE_DIRECTORY_FILE
                if directory
                else CreateOptions.FILE_NON_DIRECTORY_FILE
            ),
        )
    except Exception as exc:
        if _is_not_found(exc):
            return None
        raise
    try:
        return int(fh.end_of_file or 0)
    finally:
        fh.close()


def _mkdir_remote(tree, path: str) -> None:
    """在共享内创建单个目录（父目录必须已存在）。

    Args:
        tree: TreeConnect 对象
        path: 共享内相对目录路径

    Raises:
        Exception: 创建失败且不是「已存在」时原样抛出
    """
    from smbprotocol.open import (
        CreateDisposition,
        CreateOptions,
        DirectoryAccessMask,
        FileAttributes,
        ImpersonationLevel,
        Open,
        ShareAccess,
    )

    fh = Open(tree, path)
    try:
        fh.create(
            impersonation_level=ImpersonationLevel.Impersonation,
            desired_access=DirectoryAccessMask.FILE_READ_ATTRIBUTES,
            file_attributes=FileAttributes.FILE_ATTRIBUTE_DIRECTORY,
            share_access=(
                ShareAccess.FILE_SHARE_READ
                | ShareAccess.FILE_SHARE_WRITE
                | ShareAccess.FILE_SHARE_DELETE
            ),
            create_disposition=CreateDisposition.FILE_CREATE,
            create_options=CreateOptions.FILE_DIRECTORY_FILE,
        )
        logger.info("[SMB] 创建远端目录: %s", path)
    except Exception as exc:
        # 目录已存在（如并发创建）视为正常
        if not _is_already_exists(exc):
            raise
        logger.debug("[SMB] 远端目录已存在: %s", path)
    finally:
        fh.close()


def _ensure_remote_dir(tree, remote_dir: str) -> str:
    """确保共享内的目录存在，不存在则逐级创建。

    Args:
        tree: TreeConnect 对象
        remote_dir: 共享内相对目录（空串表示共享根目录）

    Returns:
        规范化后的远端目录路径

    Raises:
        Exception: 路径已存在但不是目录等错误原样抛出
    """
    parts = [p for p in remote_dir.replace("\\", "/").split("/") if p]
    cur = ""
    for part in parts:
        cur = f"{cur}/{part}" if cur else part
        if _open_file_stat(tree, cur, directory=True) is None:
            _mkdir_remote(tree, cur)
    return cur


def _write_file(tree, local_path: Path, remote_path: str) -> None:
    """把本地文件写入共享内指定路径（覆盖写，并回读校验大小）。

    Args:
        tree: TreeConnect 对象
        local_path: 本地文件路径
        remote_path: 远端完整路径（共享内相对路径，含文件名）

    Raises:
        RuntimeError: 写入失败或大小校验不通过
    """
    from smbprotocol.open import (
        CreateDisposition,
        CreateOptions,
        FileAttributes,
        FilePipePrinterAccessMask,
        ImpersonationLevel,
        Open,
        ShareAccess,
    )

    logger.info(
        "[SMB] 写入 %s -> %s (%.0f KB)",
        local_path.name, remote_path,
        local_path.stat().st_size / 1024,
    )

    remote_dir = remote_path.rsplit("/", 1)[0] if "/" in remote_path else ""
    if remote_dir:
        _ensure_remote_dir(tree, remote_dir)

    fh = Open(tree, remote_path)
    try:
        fh.create(
            impersonation_level=ImpersonationLevel.Impersonation,
            desired_access=FilePipePrinterAccessMask.FILE_WRITE_DATA,
            file_attributes=FileAttributes.FILE_ATTRIBUTE_NORMAL,
            share_access=ShareAccess.FILE_SHARE_READ,
            create_disposition=CreateDisposition.FILE_OVERWRITE_IF,
            create_options=CreateOptions.FILE_NON_DIRECTORY_FILE,
        )
        fh.write(local_path.read_bytes())
        fh.flush()
    except Exception as exc:
        raise RuntimeError(
            f"SMB 上传失败 ({local_path.name} -> {remote_path}): "
            f"{type(exc).__name__}: {exc}"
        ) from exc
    finally:
        fh.close()

    remote_size = _open_file_stat(tree, remote_path, directory=False)
    local_size = local_path.stat().st_size
    if remote_size != local_size:
        raise RuntimeError(
            f"SMB 上传校验失败 ({remote_path}): "
            f"远端 {remote_size} 字节 != 本地 {local_size} 字节"
        )
    logger.info("[SMB] 上传完成: %s (%d 字节)", remote_path, remote_size)


def upload_file(local_path: Path, remote_path: str) -> str:
    """把本地文件上传到 SMB 指定路径（自动创建远端目录）。

    不做存在性检查，直接覆盖写入；「不重复上传」的逻辑由
    upload_rate_file 负责。

    Args:
        local_path: 本地文件路径
        remote_path: 远端完整路径（共享内相对路径，含文件名）

    Returns:
        实际写入的远端路径

    Raises:
        RuntimeError: SMB 未配置、连接失败或上传失败
        FileNotFoundError: 本地文件不存在
    """
    local_path = Path(local_path)
    if not local_path.exists():
        raise FileNotFoundError(f"本地文件不存在: {local_path}")
    remote_path = remote_path.replace("\\", "/")

    with _open_tree() as tree:
        _write_file(tree, local_path, remote_path)
    return remote_path


def upload_rate_file(
    file_path: Path,
    date_str: str,
    remote_base: str | None = None,
) -> str | None:
    """把汇率 Excel 上传到 SMB 的 <根目录>/<YYYY>/<YYYYMMDD>/ 下（不重复上传）。

    远端已存在同名文件且字节数与本地一致时跳过，不重复上传；
    字节数不一致时重新上传。

    Args:
        file_path: 本地 Excel 文件路径，远端沿用其文件名
        date_str: 目标日期 YYYYMMDD，用作 SMB 上的子目录名（按年再分一层）
        remote_base: SMB 根目录，None 时取环境变量 SMB_REMOTE_DIR（默认空，
            表示共享根目录）

    Returns:
        实际写入的远端完整路径；远端已有同内容文件而跳过时返回 None

    Raises:
        RuntimeError: SMB 未配置、连接失败或上传失败
        FileNotFoundError: 本地文件不存在
    """
    local_path = Path(file_path)
    if not local_path.exists():
        raise FileNotFoundError(f"本地文件不存在: {local_path}")

    base = (remote_base if remote_base is not None else _smb_remote_dir())
    base = base.replace("\\", "/").strip("/")
    # 目录层级：<根目录>/<YYYY>/<YYYYMMDD>/<文件名>
    remote_path = "/".join(
        p for p in (base, date_str[:4], date_str, local_path.name) if p
    )

    local_size = local_path.stat().st_size
    with _open_tree() as tree:
        remote_size = _open_file_stat(tree, remote_path, directory=False)
        if remote_size is not None:
            if remote_size == local_size:
                logger.info(
                    "[SMB] 远端已存在同大小文件，跳过上传: %s (%d 字节)",
                    remote_path, remote_size,
                )
                return None
            logger.warning(
                "[SMB] 远端已存在但大小不一致(远端 %d != 本地 %d)，重新上传: %s",
                remote_size, local_size, remote_path,
            )
        _write_file(tree, local_path, remote_path)
    return remote_path
