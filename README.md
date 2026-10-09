# FX 汇率邮件

从中銀香港（BOCHK）与 HKEx 抓取汇率，按 TFISF Excel 公式计算自定义汇率，
生成 xlsx 附件并通过 Microsoft Graph 发送汇率报告邮件。

## 用法

```bash
uv run fxmail                       # 发送今日汇率邮件
uv run fxmail --date 20260930       # 发送指定日期汇率邮件
uv run fxmail --no-reuse            # 忽略当日缓存，强制重新抓取 BOCHK
uv run fxmail --log-level DEBUG     # 排障：输出调试级别日志
uv run fxmail --sftp-tree           # 只列出 SFTP 根目录第一层后退出（不发邮件）
uv run fxmail --sftp-tree 20261009 --sftp-depth 2
```

## SFTP 上传

给 BOC_RECIEVER 发「仅自定义汇率」邮件时，同时把该 Excel 上传到 SFTP：

```
<SFTP_DIR>/<YYYYMMDD>/自定义汇率<yymmdd>.xlsx
```

例如 `FX_Rates/20261009/自定义汇率261009.xlsx`（目录不存在会自动创建）。

| 环境变量 | 说明 |
| --- | --- |
| `SFTP` | SFTP 地址 `host[:port]`，如 `10.202.5.244:2022`（端口默认 22） |
| `SFTP_DIR` | SFTP 根目录，如 `FX_Rates` |
| `SFTP_USER` | 登录用户名 |
| `SFTP_PWD` | 登录密码 |

未配置 `SFTP` / `SFTP_USER` 时跳过上传；上传失败只记日志并发企业微信通知，不影响已发出的邮件。

### 查看 SFTP 目录树

```bash
uv run fxmail --sftp-tree                     # 整个 SFTP 根目录的第一层
uv run fxmail --sftp-tree FX_Rates            # 只看 FX_Rates 这一层
uv run fxmail --sftp-tree /FX_Rates --sftp-depth 2   # 需要更深层时显式指定
```

不带 PATH 时从整个 SFTP 根目录（登录后的默认目录，通常为 `/`）开始，**不套用 `SFTP_DIR`**；
默认层级由 `sftp_upload.DEFAULT_TREE_DEPTH` 决定（当前 4），需要更深/更浅时用 `--sftp-depth N` 覆盖。
`--sftp-tree` 只连 SFTP 打印目录后即退出，不抓汇率也不发邮件；文件名后附带大小与修改时间。

## 日志

所有模块统一通过 `log_setup.py` 输出日志：

- 控制台（stdout）+ 文件 `logs/fx_YYYYMMDD.log`，UTF-8，单文件 5MB 轮转、保留 10 份
- 第三方库（requests / urllib3 / playwright 等）默认只输出 WARNING 以上

可用环境变量或命令行参数调整：

| 配置 | 环境变量 | 命令行参数 | 默认值 |
| --- | --- | --- | --- |
| 日志级别 | `LOG_LEVEL` | `--log-level` | `INFO` |
| 日志目录 | `LOG_DIR` | `--log-dir` | `./logs` |
| 文件名前缀 | `LOG_FILE_PREFIX` | - | `fx` |

在代码里使用：

```python
from log_setup import get_logger

logger = get_logger(__name__)
logger.info("...")
```
