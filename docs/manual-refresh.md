# 在普通 ChatGPT 里手动刷新本地 RAG

仅在用户明确需要时启用。刷新只读取 Outlook、更新 Windows 本地索引，不改邮件、
已读状态或日历，不发邮件。与发送分别开关，默认 STDIO 仍是五个只读工具。

在仓库根目录、当前普通权限 Windows 用户下执行：

```powershell
# 后台运行：保留逐封确认发送，并开启手动刷新
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\autostart.ps1 -Action Install -EnableSend -EnableRefresh
# 不需要发送时，仅用 -EnableRefresh。未传入的能力会关闭。
# 前台方式（不要与后台任务同时运行）：
# .\.venv\Scripts\python.exe scripts\setup_tunnel.py run --enable-send --enable-refresh
```

其他 STDIO 客户端只需给该进程配置 `OUTLOOK_MCP_ENABLE_REFRESH=1` 并重载。
不修改全局环境。后台 Restart 保留安装时能力；关闭刷新则重新 Install，不带
`-EnableRefresh`，仍需发送则保留 `-EnableSend`。不新增任何定时同步。

重启本项目后台任务后，在普通 ChatGPT 的 OutlookMCP Windows 管理页面点击
**Refresh tools**，确认新增 `refresh_email_index` 和 `get_refresh_status`，重新加载
选中了插件的聊天。保持 **ask_before_writes**；刷新可能出现 ChatGPT 原生写操作
审批，不要改成 Always allow。发送仍须完整预览和用户后续明确确认。

直接对普通 Chat 说：**“同步一下学校邮箱到本地 RAG，并报告结果；不要发送邮件。”**

## 工具流程和固定上限

1. `refresh_email_index()` 不接收任意路径、文件夹、命令、全量重建或模型下载参数。
   立即返回 job_id；重复调用返回已有活动任务。
2. 固定读取最近 **72 小时**的 **InboxAll**（含收件箱子文件夹）和 **Sent**，每个根
   文件夹最多 **2 页 × 100 封**，复用原生同步及附件格式/大小限制。重扫近期邮件以
   检查正文/附件变化，重复同步不重复建块。
3. 本地模型存在时，最多补 **64 个近期待处理 chunk 的向量**，优先新建块。后续手动
   刷新可继续近期积压，不会处理全部历史积压或安装模型。
4. 按建议间隔调用 `get_refresh_status(job_id)`，直到 completed / partial / failed /
   timed_out / interrupted。queued 和 running 不是完成。结果包括扫描条数、附件
   状态计数、前后邮件/chunk/向量数、近期待补向量数和时间。
5. 最大 **10 分钟**；超时终止该任务及其提取/模型子进程。同步和向量阶段共用
   writer lock，不与本地 CLI 写任务重叠。中断/失败时早先已提交页面可能保留，
   不代表整次回滚；不会自动重跑，更不会重试发送。

`has_more_mail=true`、模型缺失、嵌入失败或近期向量未补完时报告 **partial**。
不支持、无文本（扫描 PDF 不做 OCR）、加密和超限附件按类别报告，不能说所有
附件都已提取。即使 completed，也只表示固定近期范围内完成，不代表全历史最新。
超过 72 小时没刷新可能留下时间缺口，需另用本地 CLI 明确回填。索引保留历史，
不会镜像 Outlook 中的删除或移动。

Windows/后台中断后报告 interrupted，不偷偷重启刷新。状态在
`data/refresh-jobs/`，只含计数、时间、错误类别等运行信息，不含邮件正文或密钥，
最多保留 30 次，目录被 Git 忽略。默认数据目录已限制当前用户和 SYSTEM；自定义
数据目录应按 INSTALL.md 单独保护。所有依赖、数据库和模型仍在 Windows，不依赖 WSL。
