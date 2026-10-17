# 完整安装指南

本指南同时面向人和负责安装的 ChatGPT。安装必须在拥有本机终端权限的 Windows 会话中执行；普通网页聊天无法仅凭仓库链接替用户执行本地安装。

## 1. 安装前检查

- Windows 10/11；classic Outlook 已安装、已完成账户登录。首次登录和 MFA 由用户在 Outlook 中完成，不索取密码。
- 以运行 Outlook 的同一 Windows 用户、普通权限启动终端；不要以管理员方式运行 Outlook。
- Windows 原生 Python 3.12–3.14（含 `py` launcher），Node.js 22 LTS 或更新版本，以及 Git。
- 如果缺少程序，可使用系统已有的软件管理器安装官方发行版；安装系统级软件需要遵守本机会话权限。然后重新打开终端，验证 `py -3 --version`、`node --version`、`git --version`。
- 建议本地 SSD 预留至少 15–20 GB；旧大型索引导入期间还需要旧库和新库同时存在。不要把活动 SQLite 放在 OneDrive、网络共享或同步盘。
- 仓库默认 `E:\wanjizheng\MyProject\OutlookMCP`，可以选择其他 Windows 路径。所有环境和依赖均安装在 checkout 内；不使用 WSL Python，也不全局安装 Python/npm 依赖。

## 2. 获取仓库并安装

执行 `git clone https://github.com/wanjizheng/OutlookMCP.git E:\wanjizheng\MyProject\OutlookMCP`，然后：

```powershell
cd E:\wanjizheng\MyProject\OutlookMCP
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -WithEmbeddings -RegisterClient
```

脚本创建 `.venv`，按锁文件安装 Python 和 Node 依赖，下载并核对本地 embedding 模型 SHA-256，保护数据目录，执行离线回归测试，检查 Outlook，然后注册 `outlook_local` MCP。它不修改账户登录、不发送邮件、不改日历、不启动 WSL。

模型约 329 MB，使用 embeddinggemma GGUF；使用须符合[模型仓库及其许可说明](https://huggingface.co/ggml-org/embeddinggemma-300m-qat-q8_0-GGUF)。已存在且哈希正确的模型不会重复下载。基础安装可省略 `-WithEmbeddings`，届时只有词法检索，hybrid 会明确报告回退。

`-RegisterClient` 只追加 `mcp_servers.outlook_local`；其他 TOML 字节保留，原配置备份放在被 Git 忽略的 `artifacts/`。若已经存在同名配置，安装脚本停止并要求检查，不会覆盖。重装可省略 `-RegisterClient`。

## 3. 建立本地邮件索引

新安装无索引时：

```powershell
.\.venv\Scripts\python.exe -m outlook_mcp.cli sync --folders InboxAll,Sent --hours 72 --pages 10
.\.venv\Scripts\python.exe -m outlook_mcp.cli status
```

`--hours 0` 表示不限时间；每页 100 封；`--pages` 限制本次调用的工作量。返回 `has_more=true` 代表尚未覆盖全部范围，不能称为完整索引。扩大 pages 或使用断点续传模式（见 CLI help）。同步只读取 Outlook 并写本地索引，不修改邮箱。

全量回填可分批执行下面的命令，`--resume` 从已经提交的页继续；反复执行直到两个文件夹都返回 `has_more=false`：

```powershell
.\.venv\Scripts\python.exe -m outlook_mcp.cli sync --folders InboxAll,Sent --hours 0 --pages 20 --resume
.\.venv\Scripts\python.exe -m outlook_mcp.cli embed-pending --limit 200
```

`embed-pending` 为原生新增/更新的文本块补本地向量，不读取 Outlook。重复运行可补齐剩余向量；不安装模型时仍可使用 BM25。

首次全量索引建议在普通终端中运行，保持终端开启；MCP 查询不自动启动全量同步。新同步正文最多 100,000 字符，附件正文保留导入的已有索引，原生附件增量支持情况见架构说明。

已有 SQLite 邮件 RAG：使用 SQLite online backup，不能只复制正在使用的 `.db` 而忽略 WAL。迁移命令必须在**源数据库所属操作系统**运行，避免跨系统锁失效。源为 Windows 时：

```powershell
.\.venv\Scripts\python.exe scripts\import_legacy.py --source C:\path\email_rag.db --destination E:\wanjizheng\MyProject\OutlookMCP\data\email_rag.db
```

源为 WSL 时，允许**仅迁移这一次**调用 WSL：

```powershell
wsl.exe -d Ubuntu -- python3 /mnt/e/wanjizheng/MyProject/OutlookMCP/scripts/import_legacy.py --source /home/YOUR_USER/.openclaw/workspace/data/email_rag/email_rag.db --destination /mnt/e/wanjizheng/MyProject/OutlookMCP/data/email_rag.db
```

导入不会改源库，不覆盖已有目标。完成后校验数据库并保留文本块和向量；此后 ChatGPT 使用新库不再访问 WSL。旧库保留给 OpenClaw，不自动删除。不同 embedding 模型即使维数相同也不可混用，导入时必须确认模型与此项目一致，否则先使用 BM25 或重建向量。

推荐将一致性备份先保存为 `data\legacy-snapshot.db`，然后用下面命令生成精简的正式库。转换会保留邮件、附件检索文本和紧凑向量，重建全文索引，去掉旧后端使用的重复 JSON 向量缓存以及无对应文本块的向量，最后执行完整 SQLite quick_check：

```powershell
.\.venv\Scripts\python.exe -m scripts.compact_legacy --source data\legacy-snapshot.db --destination data\email_rag.db
```

两种导入方式选一种，不能覆盖已经存在的正式库。精简库验收通过后才可清理本次生成的中间快照，不删除 OpenClaw 原库。

## 4. 验证与客户端接入

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe -m outlook_mcp.cli doctor
.\.venv\Scripts\python.exe scripts\smoke_mcp.py
```

smoke test 使用真正的 MCP STDIO 客户端完成初始化、列工具、读取状态、邮件、日历、RAG 检索及拒绝不存在的写工具；只输出计数或通过状态，不打印邮件正文。

当前桌面客户端使用 `%USERPROFILE%\.codex\config.toml`。若本机客户端使用其他 host/config，通过 `scripts/register_client.py --config <路径>` 注册到实际文件。或在客户端 Settings → MCP servers → Add server 选择 STDIO：

- 名称：`outlook_local`
- Command：`<checkout>\.venv\Scripts\python.exe`
- Args：`-m outlook_mcp.server`
- Working directory：`<checkout>`

重载该 MCP 连接或新开一个**本地**聊天，在 `/mcp` 或工具列表确认五个工具。然后让 ChatGPT 调用 `outlook_status`，再查询一封邮件和日历。如果当前聊天不能热加载，报告“配置和协议已验证，等待客户端重载”；不能把改配置当成已经在聊天内调用成功。不要强制关闭用户正在工作的客户端。

[官方配置与支持边界](https://learn.chatgpt.com/docs/extend/mcp)：网页版不会读取本机 config.toml；本指南安装的是桌面本地 MCP，不是网页远程 connector。

## 5. 更新、备份与卸载

- 更新前检查工作区；拉取仓库后重跑安装（注册已存在时省略 `-RegisterClient`），再运行测试。不要覆盖用户自己的修改。
- 索引备份使用 `scripts/import_legacy.py` 对新目标执行一致性备份。数据库、模型、日志、邮件、客户端配置备份永不提交 GitHub。
- 移动数据到其他 SSD：停止本项目同步，在新目录执行一致性备份，给 MCP 配置添加 `OUTLOOK_MCP_DATA_DIR` 环境变量并重载，再核对状态。代码可以留在原目录。
- 卸载时仅移除自己添加的 `mcp_servers.outlook_local` 表并重载。保留 `data/` 和 `models/`，删除它们需另行明确授权；不要恢复整个旧 config，因为它可能包含安装后其他更改。

## 6. 常见故障

- COM 无法连接：确认是 classic Outlook、同一交互用户、同一普通权限；完成 Outlook 首次运行界面。不要尝试 Graph 登录或要求学校批准 OAuth 作为隐式替代。
- 无索引：先执行同步或显式提供旧库；状态检查不会擅自创建或同步数据。
- hybrid 回退：检查模型哈希、Node 依赖安装和本地推理；可用 `search "关键词" --mode vector` 单独检查语义路径。
- 工具不出现：确认注册的是桌面当前 host 的配置、配置 TOML 可解析，以及连接已重载；查看本地日志，不输出完整配置或凭据。
- 已安装的 OpenClaw：兼容修复需独立部署 `compat/` 中的白名单文件；正常 Windows 安装不触碰它。
