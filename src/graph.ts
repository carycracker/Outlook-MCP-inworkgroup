// Microsoft Graph REST 封装
//
// 本文件是服务器与外部世界交互的唯一出口，只会访问 graph.microsoft.com。
// 没有任何动态端点拼接、没有文件写入、没有 child_process。
import type { OutlookAuth } from "./auth.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// 列表/搜索只取元数据，刻意不含 body —— 列 20 封邮件不该把正文全灌进上下文。
const SUMMARY_FIELDS = [
  "id",
  "subject",
  "from",
  "toRecipients",
  "receivedDateTime",
  "isRead",
  "hasAttachments",
  "importance",
  "bodyPreview",
  "conversationId",
  "webLink",
].join(",");

const DETAIL_FIELDS = `${SUMMARY_FIELDS},body,ccRecipients,replyTo`;

export interface EmailAddress {
  name?: string;
  address?: string;
}

interface Recipient {
  emailAddress?: EmailAddress;
}

export interface GraphMessage {
  id: string;
  subject?: string | null;
  from?: Recipient | null;
  toRecipients?: Recipient[] | null;
  ccRecipients?: Recipient[] | null;
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  importance?: string;
  bodyPreview?: string;
  body?: { contentType?: string; content?: string } | null;
  conversationId?: string;
  webLink?: string;
}

export interface GraphEvent {
  id: string;
  subject?: string | null;
  isAllDay?: boolean;
  start?: { dateTime?: string; timeZone?: string } | null;
  end?: { dateTime?: string; timeZone?: string } | null;
  location?: { displayName?: string } | null;
  organizer?: { emailAddress?: EmailAddress } | null;
  webLink?: string;
}

export interface ListOptions {
  limit: number;
  folder: string;
  unreadOnly: boolean;
}

export class GraphClient {
  constructor(private readonly auth: OutlookAuth) {}

  // 统一请求出口。
  private async request<T>(
    endpoint: string,
    options: {
      method?: string;
      body?: unknown;
      headers?: Record<string, string>;
    } = {},
  ): Promise<T> {
    const token = await this.auth.getAccessToken();
    const response = await fetch(`${GRAPH_BASE}${endpoint}`, {
      method: options.method ?? "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(options.headers ?? {}),
      },
      body:
        options.body === undefined ? undefined : JSON.stringify(options.body),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new Error(
        `Graph ${response.status} ${response.statusText} — ${detail.slice(0, 400)}`,
      );
    }
    if (response.status === 204) {
      return undefined as T;
    }
    return (await response.json()) as T;
  }

  /** 列出某文件夹中的最近邮件（仅元数据）。 */
  async listMessages(options: ListOptions): Promise<GraphMessage[]> {
    const params = [
      `$top=${options.limit}`,
      `$select=${encodeURIComponent(SUMMARY_FIELDS)}`,
      // 空格必须编码为 %20，Graph 不接受裸空格
      `$orderby=receivedDateTime%20desc`,
    ];
    if (options.unreadOnly) {
      params.push(`$filter=isRead%20eq%20false`);
    }
    const endpoint =
      `/me/mailFolders/${encodeURIComponent(options.folder)}/messages?` +
      params.join("&");

    const data = await this.request<{ value?: GraphMessage[] }>(endpoint);
    return data.value ?? [];
  }

  /** 读取单封邮件，含正文（超长截断）。 */
  async getMessage(id: string, maxBodyChars: number): Promise<GraphMessage> {
    const endpoint =
      `/me/messages/${encodeURIComponent(id)}` +
      `?$select=${encodeURIComponent(DETAIL_FIELDS)}`;

    const message = await this.request<GraphMessage>(endpoint);
    const content = message.body?.content;
    if (typeof content === "string" && content.length > maxBodyChars) {
      message.body = {
        ...message.body,
        content:
          content.slice(0, maxBodyChars) +
          `\n\n[正文已截断：原文共 ${content.length} 字符，此处显示前 ${maxBodyChars} 字符]`,
      };
    }
    return message;
  }

  /**
   * 搜索邮件。
   * Graph 不允许 $search 与 $orderby 同时出现，因此这里不排序。
   */
  async searchMessages(query: string, limit: number): Promise<GraphMessage[]> {
    const safeQuery = query.replace(/"/g, "").trim();
    const params = [
      `$search=${encodeURIComponent(`"${safeQuery}"`)}`,
      `$top=${limit}`,
      `$select=${encodeURIComponent(SUMMARY_FIELDS)}`,
    ];
    const data = await this.request<{ value?: GraphMessage[] }>(
      `/me/messages?${params.join("&")}`,
    );
    return data.value ?? [];
  }

  /** 标记已读 / 未读。 */
  async setRead(id: string, isRead: boolean): Promise<void> {
    await this.request(`/me/messages/${encodeURIComponent(id)}`, {
      method: "PATCH",
      body: { isRead },
    });
  }

  /** 新建草稿。不发送。 */
  async createDraft(options: {
    subject: string;
    body: string;
    to: string[];
  }): Promise<GraphMessage> {
    return this.request<GraphMessage>("/me/messages", {
      method: "POST",
      body: {
        subject: options.subject,
        body: { contentType: "Text", content: options.body },
        toRecipients: options.to.map((address) => ({
          emailAddress: { address },
        })),
      },
    });
  }

  /** 基于某封邮件生成回复草稿。不发送。 */
  async createReplyDraft(
    messageId: string,
    comment: string,
  ): Promise<GraphMessage> {
    return this.request<GraphMessage>(
      `/me/messages/${encodeURIComponent(messageId)}/createReply`,
      { method: "POST", body: { comment } },
    );
  }

  /**
   * 发送一封已存在的草稿。
   * 这是整个服务器唯一会对外发信的操作，默认不注册对应工具。
   */
  async sendDraft(draftId: string): Promise<void> {
    await this.request(`/me/messages/${encodeURIComponent(draftId)}/send`, {
      method: "POST",
    });
  }

  /** 读取时间段内的日历事件。 */
  async listCalendarEvents(options: {
    startDateTime: string;
    endDateTime: string;
    limit: number;
    timeZone: string;
  }): Promise<GraphEvent[]> {
    const params = [
      `startDateTime=${encodeURIComponent(options.startDateTime)}`,
      `endDateTime=${encodeURIComponent(options.endDateTime)}`,
      `$top=${options.limit}`,
      `$select=${encodeURIComponent(
        "id,subject,isAllDay,start,end,location,organizer,webLink",
      )}`,
      `$orderby=${encodeURIComponent("start/dateTime")}`,
    ];
    const data = await this.request<{ value?: GraphEvent[] }>(
      `/me/calendarView?${params.join("&")}`,
      // 让 Graph 以指定时区返回时间，省去手工换算
      { headers: { Prefer: `outlook.timezone="${options.timeZone}"` } },
    );
    return data.value ?? [];
  }

  /** 读取当前登录用户的基本资料，用于连通性自检。 */
  async getCurrentUser(): Promise<{
    displayName?: string;
    mail?: string;
    userPrincipalName?: string;
  }> {
    return this.request("/me?$select=displayName,mail,userPrincipalName");
  }
}
