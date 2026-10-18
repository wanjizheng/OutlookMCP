# 在普通 ChatGPT 里更新或全量同步 RAG

所有代码、依赖、数据库和本地模型运行在 Windows，不需要启动 WSL。
只读取 Outlook；不改正文、已读状态、文件夹、邮件或日历。发送能力独立，仍须
完整预览后，在后续用户消息中逐封明确确认。同步验收不发送邮件。

## 安装与更新

在仓库根目录，用当前普通 Windows 用户运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\autostart.ps1 -Action Install -EnableSend -EnableRefresh
```

不需要发送则只保留 `-EnableRefresh`。这些开关分别控制能力，未传入的能力会关闭。
已有后台安装保留能力并重载代码：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\autostart.ps1 -Action Restart
```

前台方式（不要与后台同时运行）：

```powershell
.\.venv\Scripts\python.exe scripts\setup_tunnel.py run --enable-send --enable-refresh
```

其他 STDIO 客户端为进程设置 `OUTLOOK_MCP_ENABLE_REFRESH=1` 后重载即可。
默认未启用时仍只有原来五个只读工具；OpenClaw 原有脚本参数不变。
在 ChatGPT 的 OutlookMCP Windows 管理页面点击 **Refresh tools**，确认出现
`refresh_email_index`、`full_sync_email_index` 和 `get_refresh_status`，然后开启
选中插件的普通 Chat。保持 ask_before_writes，不选择 Always allow。

## 对 ChatGPT 说什么

- **“更新学校邮箱 RAG，完成后报告进度和结果，不发送邮件。”**
  调用 `refresh_email_index()`。
- **“全量同步学校邮箱 RAG，完整核对后再清理本地已删除邮件的索引，不修改邮箱。”**
  调用 `full_sync_email_index()`。

两个工具都立即返回 job_id；用 `get_refresh_status(job_id)` 按建议间隔查询。
同一时刻只允许一个后台同步任务；重复请求返回活动任务及其实际 mode，不会并发。
全量请求若碰到增量正在运行，应先等待结束，再明确调用全量工具。

## 增量和成功进度

不再采用固定 72 小时窗口，也不拿最后一封邮件的 ReceivedTime 当水位线。
`native_successful_sync_at` 是整个范围成功核对的起始边界；正文和附件缓存、
每封邮件的源版本在同一个 SQLite 事务中提交。失败、超时、Windows 重启后，
再次明确说“更新 RAG”会复用已提交的版本，从未完成处补齐；不自动重试。
迁移时只接受旧刷新任务中两类文件夹均完成、无更多分页的 completed 记录作为
增量起点；它不证明全历史完整。没有可信成功记录时，先建立当前 Outlook 可见范围
的基线。旧缓存未覆盖且早于成功起点的历史记录交给全量同步，单独报告
`deferred_history`；不会因此把历史正文标成已提取。

为了捕获迟到的旧邮件、改动和移动，每次会检查所有目标文件夹的轻量元数据，
只有新增或源版本变化的邮件才重读正文及附件。即使几天没有更新，也不会受
72 小时或 200 封限制。首次迁移旧缓存时，只有源修改时间早于缓存提取时间、
主题相同且已有附件缓存的记录可以复用，否则重新读取。

## 全量、移动与删除保护

范围为 Outlook **默认 Exchange 主邮箱**所有正常邮件文件夹及子文件夹，包含
收件、已发送和主邮箱内的归档文件夹；排除 Deleted/Junk/Drafts/Outbox 和虚拟搜索
文件夹。不包括共享邮箱或单独的在线归档 Store。切换主邮箱会拒绝同步，需本地
审查范围。现有历史导入缓存按本项目原来的主邮箱缓存迁移；不要把多个邮箱的
数据库混合后交给镜像清理。

同 ID 移动只更新位置并保留块及向量。移动后 ID 改变则先读入新 ID，完整核对后
才允许清理旧 ID；不按主题合并邮件，同主题副本分别保留。增量不清理缺失 ID。

全量必须满足：所有文件夹扫描成功、所有待处理正文及附件完成、第二次完整
盘点一致、服务器历史可确认完整。存在删除候选时，先备份整个 SQLite 到
`data/sync-backups/<UTC时间>.sqlite`，再进行一次完整盘点，最后单事务清理
emails、chunks、FTS、vectors、附件缓存和源版本记录。任何读取失败或范围、
邮件变化都会阻止清理及成功水位线推进。备份含私人邮件，受数据目录权限保护，
不提交 Git；每次实际清理保留一份完整数据库，注意磁盘空间。恢复时先停止索引写
任务，由本地管理会话通过 SQLite backup API 恢复；不要直接覆盖仍在运行的 WAL
数据库。本工具不自动删除备份。

**缓存模式特别限制：** Outlook 显示“已连接”不等于所有服务器历史都在缓存中。
当前实现只在 Exchange 在线模式（ExchangeConnectionMode=800）确认完整历史。
缓存模式下全量先完成相同的增量补齐，保留历史缺口待服务器核对，但返回 partial，
`cleanup_blocked_reason=cached_exchange_cannot_prove_server_absence`、删除数量 0，
不更新 `native_full_sync_at`，保留旧历史索引。不要把该状态说成全量成功。
本程序不会更改 Outlook 缓存/账户设置。若确需服务器历史删除核对，应由用户
自行在经典 Outlook 账户设置中选择在线模式并重启 Outlook，再重试全量同步；
这可能影响离线可用性，且需学校账户策略允许。仅将缓存滑块改为“全部”并不能
通过本实现的在线完整性检查。

Outlook COM 不提供跨整个邮箱的原子快照。多次一致盘点用于检测可观察的变化，
不能保证最后一次扫描结束之后不发生新变化；之后的变化由下一次同步核对。

## 结果、附件和向量

queued/running 不是完成。terminal 状态为 completed / partial / failed /
timed_out / interrupted。结果分别报告邮件覆盖、处理/复用/扫描条数、删除条数、
成功边界、附件分类、前后邮件/chunk/向量和 pending_vectors。
正文覆盖成功但向量未齐时仍明确显示 partial；已经成功的邮件进度不会因此丢失。
增量或完整核对成功后，每次最多补 64 个缺失向量（含历史积压），按 16 块一批逐批提交，
每批调用最长三分钟；不自动下载模型。
全量被服务器覆盖检查阻止时跳过向量阶段，先报告覆盖限制，避免无效等待。FTS 可立即检索已提交
正文和提取成功的附件。已有 PDF/DOCX/PPTX/XLSX/CSV/TXT/MD 提取限制不变；
不支持的图片、扫描 PDF、加密或超限文件按类别报告，不宣称全部附件已提取。

每个任务最多运行 24 小时，超时终止它自己的子进程树，保留事务性进度。全程
复用本地 writer lock；不能与其他 CLI 同步/嵌入写任务并发。状态保存在
`data/refresh-jobs/`，保留最近 30 个任务，不含邮件正文或密钥。源版本和续传边界
保存在同一个 `email_rag.db`，因此重启隧道和 ChatGPT 不会丢失进度。
