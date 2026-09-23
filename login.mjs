// 一次性登录脚本
//
// 服务器采用懒加载认证（首次调用工具时才登录）。先在终端跑一次本脚本，
// 把令牌写入缓存，之后挂到 MCP 客户端就能静默复用，不必在 MCP 进程里登录。
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));

try {
  process.loadEnvFile?.(path.join(HERE, ".env"));
} catch {
  /* .env 不存在时忽略 */
}

const { OutlookAuth } = await import("./build/auth.js");

const clientId = process.env.OUTLOOK_CLIENT_ID;
const tenantId = process.env.OUTLOOK_TENANT_ID;

if (!clientId || !tenantId) {
  console.error("缺少 OUTLOOK_CLIENT_ID 或 OUTLOOK_TENANT_ID，请检查 .env");
  process.exit(1);
}

const auth = new OutlookAuth({ clientId, tenantId });

try {
  const token = await auth.getAccessToken();
  if (!token) {
    throw new Error("未取得访问令牌");
  }
  console.log("\n✅ 登录成功");
  console.log("令牌缓存: " + auth.getTokenCachePath());
} catch (err) {
  console.error("\n❌ 登录失败");
  console.error(err?.message || err);
  if (err?.errorCode) {
    console.error("错误码: " + err.errorCode);
  }
  process.exit(1);
}
