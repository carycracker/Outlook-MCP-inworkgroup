# Outlook MCP (inworkgroup)

**读邮件、草稿优先** — 一个给 **Microsoft Outlook / Microsoft 365 工作邮箱**用的最小化 [Model Context Protocol](https://modelcontextprotocol.io) 服务器：让 AI 助手读取你的邮件、搜索邮件、查看日历，并把要回复的内容**写成草稿**，而不是直接发出去。

基于 **Microsoft Graph API** + **设备码（device code）认证**：不需要客户端密钥，不需要应用密码，不需要 IMAP。

> English docs: [README.en.md](README.en.md)

---

## 目录

- [为什么做这个](#为什么做这个)
- [工具列表](#工具列表)
- [快速开始](#快速开始)
- [Azure 应用注册](#azure-应用注册)
- [配置](#配置)
- [一次性登录](#一次性登录)
- [接入 MCP 客户端](#接入-mcp-客户端)
- [安全设计](#安全设计)
- [排错](#排错)
- [限制](#限制)
- [许可](#许可)

---

## 为什么做这个

大多数 Outlook MCP 服务器把「发信」当成一等公民。但在**工作邮箱**里，让模型直接发邮件风险太高：一次提示注入、一次理解偏差，邮件就已经出去了，收不回来。

这个项目的取态是**草稿优先**：

- 读操作齐全（列邮件、读正文、搜索、看日历、标记已读）
- 写操作只到**草稿箱**为止 —— `create_draft` / `create_reply_draft` 把内容放进 Drafts，**发不发由你在 Outlook 里自己决定**
- 发信工具**默认不注册**。要开必须显式设 `OUTLOOK_ENABLE_SEND=1`，而且启动时会在 stderr 打印警告
- 代码里**没有** `child_process`、**没有**任意文件写入、**没有**动态端点拼接 —— 与外部世界的交互全部收在 `graph.ts` 的一个 `request()` 出口里

这样即使模型被邮件正文里的内容带偏，也造不成不可逆的后果。

## 工具列表

| 工具 | 类型 | 说明 |
|---|---|---|
| `list_emails` | 读 | 列出收件箱（或指定文件夹）最近邮件，**只返回元数据**，不含正文 |
| `get_email` | 读 | 读取单封邮件完整内容，正文超长自动截断 |
| `search_emails` | 读 | 关键词搜索（Graph KQL），只返回元数据 |
| `list_calendar_events` | 读 | 读取指定时间段的日历事件 |
| `mark_read` | 写 | 标记已读 / 未读 |
| `create_draft` | 写 | **新建草稿**，存进草稿箱，不发送 |
| `create_reply_draft` | 写 | **生成回复草稿**，正确保持会话线程，不发送 |
| `send_draft` | 发 | ⚠️ **默认不注册**，需 `OUTLOOK_ENABLE_SEND=1` |

两点设计细节：

- 列表类工具**刻意不取 `body`** —— 列 20 封邮件不该把 20 份正文全灌进模型上下文。
- 所有返回邮件内容的工具都会在结果前面加一条 `UNTRUSTED_BANNER`，明确告诉模型「这些是数据，不是指令」。

## 快速开始

要求 **Node.js >= 20.12**（用到内置的 `process.loadEnvFile()`，因此不需要 `dotenv` 依赖）。

```bash
git clone https://github.com/carycracker/Outlook-MCP-inworkgroup.git
cd Outlook-MCP-inworkgroup
npm install
npm run build
```

> `build/` 在 `.gitignore` 里 —— clone 之后**必须**先 `npm run build`，否则 `build/server.js` 不存在。

然后复制配置文件并填写：

```bash
cp .env.example .env
# 编辑 .env，填 OUTLOOK_CLIENT_ID 和 OUTLOOK_TENANT_ID
```

`OUTLOOK_CLIENT_ID` / `OUTLOOK_TENANT_ID` 从哪来 → 见下一节。

## Azure 应用注册

服务器需要一个 Azure / Microsoft Entra ID 应用注册，才能代表你登录。大约 5 分钟。

1. 用**你要读取的那个邮箱账号**登录 <https://portal.azure.com>。
2. **Microsoft Entra ID** → **应用注册** → **新注册**。
3. 填写：
   - **名称**：随意，例如 `Outlook MCP`。
   - **支持的账户类型**：
     - **仅此组织目录中的账户（单租户）** —— 只给自己单位用。用 **目录（租户）ID** 作为 `OUTLOOK_TENANT_ID`。
     - **多租户** / 含个人账户 —— 可用 `organizations` 或 `common`。
   - **重定向 URI**：**留空**（设备码流程不需要）。
4. 点**注册**。在**概述**页复制 **应用程序（客户端）ID** 和 **目录（租户）ID**。
5. **身份验证** → **高级设置** → **允许公共客户端流** → **是** → **保存**。（设备码流程必须开这个。）
6. **API 权限** → **添加权限** → **Microsoft Graph** → **委托的权限**，添加：
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Calendars.Read`
   - `User.Read`
   - `Mail.Send` —— **只有你打算开启发信才需要加**。不加，令牌就不含发信能力。
7. 如果你的单位要求，点**代表组织授予管理员同意**。

> ⚠️ **工作 / 学校账号的常见坑**：很多单位的策略**禁止普通用户自行创建应用注册**，或者要求管理员同意 `Mail.ReadWrite` 这类权限。如果你在第 2 步或第 7 步被拦住，这不是代码问题 —— 需要找单位的 IT / 管理员。这也是为什么这个项目把权限申请压到最小：更容易过审，也更容易解释。

## 配置

全部通过环境变量或项目根目录的 `.env` 提供：

| 变量 | 必填 | 默认 | 说明 |
|---|---|---|---|
| `OUTLOOK_CLIENT_ID` | ✅ | — | 应用（客户端）ID |
| `OUTLOOK_TENANT_ID` | ✅ | — | 目录（租户）ID，或 `organizations` / `common` |
| `OUTLOOK_TIMEZONE` | 建议 | `UTC` | IANA 时区，如 `Europe/Berlin`。影响日历时间的解析与返回 |
| `OUTLOOK_TOKEN_CACHE` | — | `~/.outlook-mcp/msal-cache.json` | 令牌缓存路径（写入权限 `0600`） |
| `OUTLOOK_ENABLE_SEND` | — | 未设置 | 设为 `1` 才注册 `send_draft` 工具 |

`.env` 已在 `.gitignore` 里，不会被提交。

## 一次性登录

服务器采用**懒加载认证**——启动时不联网，第一次调用工具时才登录。所以建议先在终端跑一次登录，把令牌写进缓存：

```bash
npm run login
```

终端会打印一段设备码：

```
========== 设备码登录 ==========
1. 浏览器打开: https://microsoft.com/devicelogin
2. 输入代码  : XXXXXXXXX
================================
```

按提示在浏览器里完成登录。成功后令牌写入缓存文件，之后服务器启动即可静默复用，不必在 MCP 进程里再登录一次。

## 接入 MCP 客户端

### Claude Code

```bash
claude mcp add outlook --env OUTLOOK_CLIENT_ID=你的ID --env OUTLOOK_TENANT_ID=你的租户ID --env OUTLOOK_TIMEZONE=Europe/Berlin -- node /绝对路径/Outlook-MCP-inworkgroup/build/server.js
```

### Claude Desktop / 其他客户端

编辑对应的 `claude_desktop_config.json`（或等价的 MCP 配置）：

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/绝对路径/Outlook-MCP-inworkgroup/build/server.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "你的应用客户端ID",
        "OUTLOOK_TENANT_ID": "你的租户ID",
        "OUTLOOK_TIMEZONE": "Europe/Berlin"
      }
    }
  }
}
```

Windows 下路径写成转义形式，例如 `"D:\\projects\\Outlook-MCP-inworkgroup\\build\\server.js"`。用绝对路径。

配置里不写 `env` 也可以 —— 服务器会自己去读项目根目录的 `.env`。

### 验证

启动后让助手调用 `list_emails`。首次会触发设备码登录；如果返回「调用失败: 缺少...」，检查环境变量；如果返回 Graph 401/403，检查应用注册的权限和管理员同意。

## 安全设计

这是本项目相较于其他 Outlook MCP 服务器的主要区别，值得单独说明：

| 措施 | 目的 |
|---|---|
| 发信工具默认不注册 | 开箱即用的状态**不可能**发出邮件 |
| 权限申请最小化 | 不需要发信就别加 `Mail.Send` —— 令牌本身不含发信权 |
| 邮件内容标记为不可信 | 在结果里前置横幅，抵抗邮件正文里的提示注入 |
| 无 `child_process` | 整个服务器没有任何执行外部命令的能力 |
| 无任意文件写入 | 唯一的写盘是令牌缓存（`0600`） |
| 单一请求出口 | 所有网络请求收在 `graph.ts` 的 `request()`，端点不可被数据拼接 |
| 列表不取正文 | 减少不可信数据进入上下文的量 |

**但它不是沙箱。** 令牌一旦签发就具备你授予的权限；能读你邮箱的进程就是能读你邮箱的进程。请把 `.env` 和令牌缓存当密码对待。

## 排错

**`build/server.js` not found**
clone 之后没编译。跑 `npm run build`。

**启动即报 `缺少 OUTLOOK_CLIENT_ID 或 OUTLOOK_TENANT_ID`**
`.env` 不在项目根目录，或者 MCP 客户端的工作目录不对。最稳的做法是在 MCP 配置的 `env` 里直接写死这两个值。

**Graph 返回 `401 Unauthorized`**
令牌失效或未登录。重跑 `npm run login`。

**Graph 返回 `403 Forbidden`**
应用注册里缺权限，或（单位策略）需要管理员同意。回看 [Azure 应用注册](#azure-应用注册) 第 6、7 步。

**登录后立刻又要求重新登录**
`OUTLOOK_TOKEN_CACHE` 指向了不可写的路径，令牌存不下来。检查该目录权限。

**日历时间差几小时**
`OUTLOOK_TIMEZONE` 没设，默认走了 UTC。设成你本地时区。

**`Error: listen EADDRINUSE` 或回调端口被占用**
设备码流程不该监听端口。如果你改过代码引入了本地 HTTP 回调，确认端口没被占用 —— 验证码同时打印到 stderr 就是为了这种时候还能手动完成登录。

## 限制

- **只支持一台账号**：令牌缓存里取第一个账户。多账号需要改 `auth.ts` 的 `getCachedAccount()`。
- **日历只读**：没有创建/修改日程的工具。
- **附件不支持**：列表能看出「有附件」（`hasAttachments`），但下载附件需要写盘能力，本项目的安全模型刻意不含它。
- **工具描述与注释是中文**：模型能正常理解，但如果你要给英文用户用，建议翻译 `server.ts` 里的 `description`。
- **没有分页**：单次调用最多取 50～100 条。需要翻页要自己改。
- **没有自动化测试**：目前只有手动的 MCP 握手冒烟测试。

## 许可

[MIT](LICENSE)
