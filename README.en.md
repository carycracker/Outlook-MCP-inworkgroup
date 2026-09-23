# Outlook MCP (inworkgroup)

**Read mail, draft first.** A minimal [Model Context Protocol](https://modelcontextprotocol.io) server for **Microsoft Outlook / Microsoft 365 work mailboxes**. It lets an AI assistant read your email, search it, check your calendar, and write what it wants to say **into your Drafts folder** — instead of sending it.

Built on the **Microsoft Graph API** with **device-code authentication**: no client secret, no app password, no IMAP.

> 中文文档：[README.md](README.md)

---

## Contents

- [Why this exists](#why-this-exists)
- [Tools](#tools)
- [Quick start](#quick-start)
- [Azure app registration](#azure-app-registration)
- [Configuration](#configuration)
- [One-time sign-in](#one-time-sign-in)
- [Wiring it into an MCP client](#wiring-it-into-an-mcp-client)
- [Security design](#security-design)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [License](#license)

---

## Why this exists

Most Outlook MCP servers treat "send mail" as a first-class feature. In a **work mailbox** that is a lot of risk: one prompt injection, one misread instruction, and the mail is already gone — you cannot recall it.

This project takes a **draft-first** stance:

- Read operations are complete (list, read, search, calendar, mark read)
- Write operations stop at **Drafts** — `create_draft` / `create_reply_draft` file the content away, and **you decide in Outlook whether it goes out**
- The send tool is **not registered by default**. Enabling it requires explicitly setting `OUTLOOK_ENABLE_SEND=1`, and the server prints a warning to stderr on startup
- No `child_process`, no arbitrary file writes, no dynamic endpoint assembly — every interaction with the outside world funnels through a single `request()` in `graph.ts`

So even if the model gets steered by something inside an email body, it cannot cause an irreversible outcome.

## Tools

| Tool | Kind | Description |
|---|---|---|
| `list_emails` | read | Recent messages in the inbox (or a folder). **Metadata only**, no bodies |
| `get_email` | read | Full message content, body truncated past a limit |
| `search_emails` | read | Keyword search (Graph KQL). Metadata only |
| `list_calendar_events` | read | Calendar events in a time range |
| `mark_read` | write | Mark a message read / unread |
| `create_draft` | write | **New draft** in Drafts. Does not send |
| `create_reply_draft` | write | **Reply draft**, threading preserved. Does not send |
| `send_draft` | send | ⚠️ **Not registered by default** — needs `OUTLOOK_ENABLE_SEND=1` |

Two deliberate details:

- List tools **omit `body` on purpose** — listing 20 messages should not dump 20 email bodies into the model's context.
- Every tool that returns mail content prefixes the result with an `UNTRUSTED_BANNER`, telling the model plainly that this is data, not instructions.

## Quick start

Requires **Node.js >= 20.12** (uses the built-in `process.loadEnvFile()`, so there is no `dotenv` dependency).

```bash
git clone https://github.com/carycracker/Outlook-MCP-inworkgroup.git
cd Outlook-MCP-inworkgroup
npm install
npm run build
```

> `build/` is gitignored — you **must** run `npm run build` after cloning, or `build/server.js` will not exist.

Then create your config:

```bash
cp .env.example .env
# edit .env and fill in OUTLOOK_CLIENT_ID and OUTLOOK_TENANT_ID
```

Where those two values come from → next section.

## Azure app registration

The server needs an Azure / Microsoft Entra ID app registration so it can sign in as you. About 5 minutes.

1. Sign in to <https://portal.azure.com> with **the account whose mailbox you want to read**.
2. **Microsoft Entra ID** → **App registrations** → **New registration**.
3. Fill in:
   - **Name:** anything, e.g. `Outlook MCP`.
   - **Supported account types:**
     - **Accounts in this organizational directory only (single tenant)** — for your org only. Use your **Directory (tenant) ID** as `OUTLOOK_TENANT_ID`.
     - **Multi-tenant** / including personal accounts — you may use `organizations` or `common`.
   - **Redirect URI:** leave **blank** (device code flow does not use one).
4. Click **Register**. On the **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**.
5. **Authentication** → **Advanced settings** → **Allow public client flows** → **Yes** → **Save**. (Required for device-code flow.)
6. **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**. Add:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Calendars.Read`
   - `User.Read`
   - `Mail.Send` — **only if you intend to enable sending.** Without it, the token carries no send capability at all.
7. If your organization requires it, click **Grant admin consent**.

> ⚠️ **The usual snag with work / school accounts:** plenty of orgs **forbid ordinary users from creating app registrations**, or require admin consent for something like `Mail.ReadWrite`. If you get blocked at step 2 or step 7, that is policy, not a code problem — you need your IT / admin. It is partly why this project keeps the requested scopes minimal: easier to get approved, easier to justify.

## Configuration

Everything comes from environment variables, or from a `.env` in the project root:

| Variable | Required | Default | Notes |
|---|---|---|---|
| `OUTLOOK_CLIENT_ID` | ✅ | — | Application (client) ID |
| `OUTLOOK_TENANT_ID` | ✅ | — | Directory (tenant) ID, or `organizations` / `common` |
| `OUTLOOK_TIMEZONE` | recommended | `UTC` | IANA timezone, e.g. `Europe/Berlin`. Affects how calendar times are interpreted and returned |
| `OUTLOOK_TOKEN_CACHE` | — | `~/.outlook-mcp/msal-cache.json` | Token cache path (written `0600`) |
| `OUTLOOK_ENABLE_SEND` | — | unset | Set to `1` to register the `send_draft` tool |

`.env` is in `.gitignore` and will not be committed.

## One-time sign-in

Authentication is **lazy** — the server does not touch the network at startup, only on the first tool call. So it is worth signing in once from the terminal to write the token into the cache:

```bash
npm run login
```

It prints a device code:

```
========== Device code sign-in ==========
1. Open in browser: https://microsoft.com/devicelogin
2. Enter code     : XXXXXXXXX
=========================================
```

Complete the sign-in in your browser. The token lands in the cache file, and the server reuses it silently from then on — no sign-in inside the MCP process.

## Wiring it into an MCP client

### Claude Code

```bash
claude mcp add outlook --env OUTLOOK_CLIENT_ID=YOUR_ID --env OUTLOOK_TENANT_ID=YOUR_TENANT_ID --env OUTLOOK_TIMEZONE=Europe/Berlin -- node /absolute/path/Outlook-MCP-inworkgroup/build/server.js
```

### Claude Desktop / other clients

Edit the relevant `claude_desktop_config.json` (or equivalent MCP config):

```json
{
  "mcpServers": {
    "outlook": {
      "command": "node",
      "args": ["/absolute/path/Outlook-MCP-inworkgroup/build/server.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-application-client-id",
        "OUTLOOK_TENANT_ID": "your-tenant-id",
        "OUTLOOK_TIMEZONE": "Europe/Berlin"
      }
    }
  }
}
```

On Windows, escape the path: `"D:\\projects\\Outlook-MCP-inworkgroup\\build\\server.js"`. Use an absolute path.

You can omit the `env` block entirely — the server also reads a `.env` from the project root.

### Verifying

Ask the assistant to call `list_emails`. The first call triggers device-code sign-in. If you get "missing OUTLOOK_CLIENT_ID", check your environment; if Graph returns 401/403, check the app registration's permissions and admin consent.

## Security design

This is the main thing that distinguishes this project from other Outlook MCP servers, so it deserves its own section:

| Measure | Purpose |
|---|---|
| Send tool unregistered by default | Out of the box it is **impossible** to send mail |
| Minimal scopes | Skip `Mail.Send` unless you need it — the token then has no send capability |
| Mail content flagged untrusted | A banner prepended to results, resisting prompt injection from email bodies |
| No `child_process` | The server has no ability to execute external commands at all |
| No arbitrary file writes | The only disk write is the token cache (`0600`) |
| Single request chokepoint | All network calls go through `request()` in `graph.ts`; endpoints cannot be assembled from data |
| No bodies in list results | Less untrusted data entering the context |

**But this is not a sandbox.** Once a token is issued it carries whatever you granted. A process that can read your mailbox is a process that can read your mailbox. Treat `.env` and the token cache like passwords.

## Troubleshooting

**`build/server.js` not found**
You cloned and did not compile. Run `npm run build`.

**Startup fails with "missing OUTLOOK_CLIENT_ID or OUTLOOK_TENANT_ID"**
`.env` is not in the project root, or the MCP client's working directory differs. Most robust fix: set both values directly in the MCP config's `env` block.

**Graph returns `401 Unauthorized`**
Token expired or never signed in. Re-run `npm run login`.

**Graph returns `403 Forbidden`**
Missing permission on the app registration, or admin consent is required by policy. Revisit steps 6–7 of [Azure app registration](#azure-app-registration).

**Asked to sign in again immediately after signing in**
`OUTLOOK_TOKEN_CACHE` points somewhere unwritable, so the token cannot be persisted. Check directory permissions.

**Calendar times are off by hours**
`OUTLOOK_TIMEZONE` is unset, so it defaulted to UTC. Set your local timezone.

**`Error: listen EADDRINUSE` / a callback port is taken**
Device-code flow should not be listening on a port. If you modified the code to add a local HTTP callback, confirm the port is free — the code being printed to stderr as well is exactly so sign-in still works when a browser helper cannot launch.

## Limitations

- **One account only:** the cache uses the first account found. Multi-account needs changes in `getCachedAccount()` in `auth.ts`.
- **Calendar is read-only:** no tools to create or modify events.
- **No attachments:** listings show `hasAttachments`, but downloading an attachment means writing to disk, which this project's security model deliberately omits.
- **Tool descriptions and comments are in Chinese:** models handle it fine, but if you are targeting English-speaking users, translate the `description` strings in `server.ts`.
- **No pagination:** a single call returns at most 50–100 items. Paging is up to you.
- **No automated tests:** currently only a manual MCP handshake smoke test.

## License

[MIT](LICENSE)
