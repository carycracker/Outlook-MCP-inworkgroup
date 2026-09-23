// 认证层：MSAL 设备码流程 + 本地令牌缓存 + 静默刷新
//
// 设计要点：
//   * 不使用 child_process / exec —— 不自动打开浏览器，验证码只打印到 stderr。
//     这样整个服务器没有任何执行外部命令的能力。
//   * 令牌缓存写在固定路径，权限 0600，可用 OUTLOOK_TOKEN_CACHE 覆盖。
//   * 申请的权限与 Azure 应用注册中已配置的委派权限一致。
import {
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
  type DeviceCodeRequest,
  type SilentFlowRequest,
} from "@azure/msal-node";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

// 申请的权限 —— 需与 Azure 应用注册中已授予的委派权限一致。
//
// 注意：这里含 Mail.Send，因此令牌具备发信能力。但服务器默认不注册发信工具，
// 需要显式设置 OUTLOOK_ENABLE_SEND=1 才会启用（见 server.ts）。
// 若你的场景完全不需要发信，可以从下面删掉 Mail.Send 这一行 —— 令牌就不含发信权限，
// 这是"代码里不注册发信工具"之外的第二道锁。
export const GRAPH_SCOPES: string[] = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.Read",
  "https://graph.microsoft.com/User.Read",
  "offline_access", // 换取刷新令牌，实现长期免登录
];

// 令牌缓存位置。默认写到用户主目录下的 .outlook-mcp/，可用 OUTLOOK_TOKEN_CACHE
// 指向别处（例如放进某个 MCP 客户端的私有目录）。
const CACHE_PATH =
  process.env.OUTLOOK_TOKEN_CACHE ||
  path.join(os.homedir(), ".outlook-mcp", "msal-cache.json");

export interface OutlookAuthConfig {
  clientId: string;
  tenantId: string;
}

export class OutlookAuth {
  private readonly pca: PublicClientApplication;

  constructor(config: OutlookAuthConfig) {
    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    };
    this.pca = new PublicClientApplication(msalConfig);
    this.loadCache();
  }

  getTokenCachePath(): string {
    return CACHE_PATH;
  }

  private loadCache(): void {
    try {
      if (fs.existsSync(CACHE_PATH)) {
        this.pca.getTokenCache().deserialize(fs.readFileSync(CACHE_PATH, "utf-8"));
      }
    } catch (error) {
      console.error(
        "无法读取令牌缓存:",
        error instanceof Error ? error.message : error,
      );
    }
  }

  private saveCache(): void {
    const dir = path.dirname(CACHE_PATH);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(CACHE_PATH, this.pca.getTokenCache().serialize(), {
      mode: 0o600,
    });
  }

  private async getCachedAccount(): Promise<AccountInfo | null> {
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    return accounts.length > 0 ? accounts[0]! : null;
  }

  private async authenticateWithDeviceCode(): Promise<string> {
    const request: DeviceCodeRequest = {
      scopes: GRAPH_SCOPES,
      deviceCodeCallback: (response) => {
        console.error("\n========== 设备码登录 ==========");
        console.error(`1. 浏览器打开: ${response.verificationUri}`);
        console.error(`2. 输入代码  : ${response.userCode}`);
        console.error("================================\n");
      },
    };

    const result = await this.pca.acquireTokenByDeviceCode(request);
    if (!result?.accessToken) {
      throw new Error("未取得访问令牌");
    }
    this.saveCache();
    console.error(`已登录: ${result.account?.username ?? "未知账号"}`);
    return result.accessToken;
  }

  /** 优先静默刷新；仅在缓存失效时才走设备码登录。 */
  async getAccessToken(): Promise<string> {
    const account = await this.getCachedAccount();
    if (account) {
      try {
        const request: SilentFlowRequest = { scopes: GRAPH_SCOPES, account };
        const result = await this.pca.acquireTokenSilent(request);
        if (result?.accessToken) {
          this.saveCache();
          return result.accessToken;
        }
      } catch {
        console.error("静默刷新失败，将重新走设备码登录");
      }
    }
    return this.authenticateWithDeviceCode();
  }
}
