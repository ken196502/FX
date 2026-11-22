# FX 汇率邮件

从中銀香港（BOCHK）与 HKEx 抓取汇率，按 TFISF Excel 公式计算自定义汇率，
生成 xlsx 附件并通过 Microsoft Graph 发送汇率报告邮件。

## 用法

```bash
uv run fxmail                       # 默认流程（见「默认流程」）
uv run fxmail --date 20260930       # 默认流程，指定日期
uv run fxmail --upload-smb 20260930 # 只上传 SMB，不发邮件
uv run fxmail --no-reuse            # 忽略当日缓存，强制重新抓取 BOCHK / HKEx
uv run fxmail --log-level DEBUG     # 排障：输出调试级别日志
uv run fxmail --sftp-tree           # 只列出 SFTP 根目录第一层后退出（不发邮件）
uv run fxmail --sftp-tree 20261009 --sftp-depth 2
```

## SFTP 上传

给 BOC_RECIEVER 发「仅自定义汇率」邮件时，同时把该 Excel 上传到 SFTP；
港元最优惠利率表 `base_rate<yymmdd>.xlsx` 每次默认流程都会同步到同一日期目录
（不受邮件收件人配置影响）：

```
<SFTP_DIR>/<YYYYMMDD>/自定义汇率<yymmdd>.xlsx
<SFTP_DIR>/<YYYYMMDD>/base_rate<yymmdd>.xlsx
```

例如 `FX_Rates/20261009/自定义汇率261009.xlsx`、`FX_Rates/20261009/base_rate261009.xlsx`
（目录不存在会自动创建）。`base_rate` 表字段为
`base_rate_code / base_rate_name / money_type / saving_rate / prime_rate /
day_year / vaild_date / remark`，其中 `saving_rate` 固定为 0，`prime_rate` 取
BOCHK 港元最优惠利率的小数值（5.00000% → 0.05），`vaild_date` 是有效期限，
固定写死为 `20991231`（与抓取日期无关）。

**远端已存在同名文件时默认跳过，不覆盖**。两个覆盖开关：

```bash
uv run fxmail --upload-sftp 20261009                    # 已存在则跳过
uv run fxmail --upload-sftp 20261009 --force            # 覆盖所有 SFTP 同名文件（含自定义汇率）
uv run fxmail --date 20261009 --force-base-rate         # 只覆盖 SFTP 上的 base_rate
```

`--force` / `--force-base-rate` 对默认流程、`--upload-smb`、`--upload-sftp` 中的 SFTP 上传都生效；SMB 上传仍按同大小跳过，不受这两个开关影响。

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

## 默认流程（uv run fxmail）

`uv run fxmail` 不带参数时执行（`--upload-smb YYYYMMDD` 是它的「只上传、不发邮件」版本）：

1. **当天还没抓取过**（`temp/` 里没有当天 Excel）：先尝试抓取——检测到 `FX_RECIEVER`（或未配置任何收件人）时**银行（BOCHK）和交易所（HKEx）都抓**；只有 `BOC_RECIEVER` 时只抓银行；两种模式都同时抓 [BOCHK 港元最优惠利率](https://www.bochk.com/en/investment/rates/hkprime.html)，在 `temp/` 生成 `base_rate<yymmdd>.xlsx`；
2. **已配置 SMB**：把当天已抓取的汇率 Excel 上传到 `<SMB_REMOTE_DIR>/<YYYY>/<YYYYMMDD>/`，**远端已存在同大小同名文件时跳过，不重复上传**；大小不一致才重新上传；
3. **按「没有上传过的才发邮件」发邮件**：
   - 检测到 `BOC_RECIEVER`：自定义汇率**上传后就发**「仅自定义汇率」邮件（并发往 SFTP）；**已经上传过（本轮跳过）就不发**；
   - 检测到 `FX_RECIEVER`：两个文件都上传（跳过已上传）后，**两个附件放一起发一封邮件**；无交易所汇率时不发，避免误导；
4. 将 `base_rate<yymmdd>.xlsx` 同步到 SFTP 的同日期目录（未配置 SFTP 时跳过，不影响邮件）；
5. **未配置 SMB**：只在本地生成 Excel，不上传；此时所有本地文件按「未上传过」处理，照常发邮件。

任一文件上传失败都会记日志、发企业微信通知并以退出码 1 结束，不做静默兜底；BOCHK 等关键数据异常时不上传、不发汇率邮件（给 `BOC_RECIEVER` 发错误报告），退出码 1。

| 环境变量 | 说明 |
| --- | --- |
| `SMB_SERVER` | 服务器地址，可写 `10.202.5.244` / `10.202.5.244/FX_Rates` / `\\10.202.5.244\FX_Rates` |
| `SMB_SHARE` | 共享名（`SMB_SERVER` 中未带共享名时必填），如 `FX_Rates` |
| `SMB_REMOTE_DIR` | 共享内的根目录，如 `FX_Rates/import`（缺省表示共享根目录） |
| `SMB_USERNAME` | 登录用户名 |
| `SMB_PASSWORD` | 登录密码 |

远端布局为 `<SMB_REMOTE_DIR>/<YYYY>/<YYYYMMDD>/<文件名>`，例如 `FX_Rates/2026/20261009/自定义汇率261009.xlsx`（目录不存在会自动创建）。

同日复用：自定义汇率、交易所汇率、港元最优惠利率与截图都进当日快照（`temp/custom_snapshot_yyyymmdd.json`），同日第二次及以后的运行不重复抓取、不重复改写 Excel（文件字节稳定，「已上传」判断才可靠）。

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
