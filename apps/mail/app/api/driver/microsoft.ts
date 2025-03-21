import { type IConfig, type MailManager } from "./types";
import { Client } from "@microsoft/microsoft-graph-client";
import { EnableBrain } from "@/actions/brain";
import { type ParsedMessage } from "@/types";
import * as he from "he";
import { parseFrom, wasSentWithTLS } from "@/lib/email-utils";

const findHtmlBody = (parts: any[]): string => {
  for (const part of parts) {
    if (part.mimeType === "text/html" && part.body?.content) {
      console.log("✓ Driver: Found HTML content in message part");
      return part.body.content;
    }
    if (part.parts) {
      const found = findHtmlBody(part.parts);
      if (found) return found;
    }
  }
  console.log("⚠️ Driver: No HTML content found in message parts");
  return "";
};

export const driver = async (config: IConfig): Promise<MailManager> => {
  const client = Client.init({
    authProvider: (done) => {
      done(null, config.auth?.access_token);
    },
  });

  const getScope = () =>
    [
      "https://graph.microsoft.com/Mail.ReadWrite",
      "https://graph.microsoft.com/User.Read",
    ].join(" ");

  if (config.auth) {
    EnableBrain()
      .then(() => console.log("✅ Driver: Enabled"))
      .catch(() => console.log("✅ Driver: Enabled"));
  }

  const parse = ({
    id,
    threadId,
    snippet,
    labelIds,
    payload,
  }: any): Omit<
    ParsedMessage,
    "body" | "processedHtml" | "blobUrl" | "totalReplies"
  > => {
    const receivedOn =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "date")?.value || "Failed";
    const sender =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "from")?.value || "Failed";
    const subject =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "subject")?.value || "";
    const references =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "references")?.value || "";
    const inReplyTo =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "in-reply-to")?.value || "";
    const messageId =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "message-id")?.value || "";
    const listUnsubscribe =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "list-unsubscribe")?.value ||
      undefined;
    const listUnsubscribePost =
      payload?.headers?.find((h: any) => h.name?.toLowerCase() === "list-unsubscribe-post")?.value ||
      undefined;

    const receivedHeaders = payload?.headers?.filter((header: any) => header.name?.toLowerCase() === 'received')
      .map((header: any) => header.value || '') || [];
    const hasTLSReport = payload?.headers?.some((header: any) => header.name?.toLowerCase() === 'tls-report');

    return {
      id: id || "ERROR",
      threadId: threadId || "",
      title: snippet ? he.decode(snippet).trim() : "ERROR",
      tls: wasSentWithTLS(receivedHeaders) || !!hasTLSReport,
      tags: labelIds || [],
      listUnsubscribe,
      listUnsubscribePost,
      references,
      inReplyTo,
      sender: parseFrom(sender),
      unread: labelIds ? labelIds.includes("UNREAD") : false,
      receivedOn,
      subject: subject ? subject.replace(/"/g, "").trim() : "(no subject)",
      messageId,
    };
  };

  const manager = {
    getAttachment: async (messageId: string, attachmentId: string) => {
      try {
        const response = await client
          .api(`/me/messages/${messageId}/attachments/${attachmentId}`)
          .get();

        const attachmentData = response.contentBytes || "";

        return attachmentData;
      } catch (error) {
        console.error("Error fetching attachment:", error);
        throw error;
      }
    },
    markAsRead: async (id: string[]) => {
      await Promise.all(
        id.map(async (messageId) => {
          await client
            .api(`/me/messages/${messageId}`)
            .update({ isRead: true });
        })
      );
    },
    markAsUnread: async (id: string[]) => {
      await Promise.all(
        id.map(async (messageId) => {
          await client
            .api(`/me/messages/${messageId}`)
            .update({ isRead: false });
        })
      );
    },
    getScope,
    getUserInfo: async (tokens: { access_token: string; refresh_token: string }) => {
      const response = await client
        .api("/me")
        .get();
      return response;
    },
    getTokens: async <T>(code: string) => {
      try {
        const response = await client
          .api("/oauth2/v2.0/token")
          .post({
            grant_type: "authorization_code",
            code,
            redirect_uri: process.env.MICROSOFT_REDIRECT_URI,
            client_id: process.env.MICROSOFT_CLIENT_ID,
            client_secret: process.env.MICROSOFT_CLIENT_SECRET,
          });
        return { tokens: response } as T;
      } catch (error) {
        console.error("Error getting tokens:", error);
        throw error;
      }
    },
    generateConnectionAuthUrl: (userId: string) => {
      const params = new URLSearchParams({
        client_id: process.env.MICROSOFT_CLIENT_ID as string,
        response_type: "code",
        redirect_uri: process.env.MICROSOFT_REDIRECT_URI as string,
        response_mode: "query",
        scope: getScope(),
        state: userId,
      });
      return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`;
    },
    count: async () => {
      const response = await client
        .api("/me/mailFolders/inbox/messages")
        .filter("isRead eq false")
        .count(true)
        .get();
      return { count: response["@odata.count"] };
    },
    list: async (
      folder: string,
      q: string,
      maxResults = 20,
      _labelIds: string[] = [],
      pageToken?: string,
    ) => {
      const response = await client
        .api(`/me/mailFolders/${folder}/messages`)
        .filter(q)
        .top(maxResults)
        .skipToken(pageToken)
        .get();

      const threads = await Promise.all(
        response.value.map(async (message: any) => {
          const parsed = parse(message);
          return {
            ...parsed,
            body: "",
            processedHtml: "",
            blobUrl: "",
            totalReplies: 0,
            threadId: message.id,
          };
        })
      );

      return { ...response, threads } as any;
    },
    get: async (id: string): Promise<ParsedMessage[]> => {
      const response = await client
        .api(`/me/messages/${id}`)
        .expand("attachments")
        .get();

      const bodyData = response.body?.content || "";

      const parsedData = parse(response);

      const attachments = response.attachments?.map((attachment: any) => ({
        filename: attachment.name || "",
        mimeType: attachment.contentType || "",
        size: Number(attachment.size || 0),
        attachmentId: attachment.id,
        headers: [],
        body: attachment.contentBytes,
      })) || [];

      const fullEmailData = {
        ...parsedData,
        body: "",
        processedHtml: "",
        blobUrl: "",
        decodedBody: bodyData,
        attachments,
      };

      return [fullEmailData];
    },
    create: async (data: any) => {
      const response = await client
        .api("/me/sendMail")
        .post({ message: data });
      return response;
    },
    delete: async (id: string) => {
      const response = await client
        .api(`/me/messages/${id}`)
        .delete();
      return response;
    },
    normalizeIds: (ids: string[]) => {
      const normalizedIds: string[] = [];
      const threadIds: string[] = [];

      for (const id of ids) {
        if (id.startsWith("thread:")) {
          threadIds.push(id.substring(7));
        } else {
          normalizedIds.push(id);
        }
      }

      return { normalizedIds, threadIds };
    },
    async modifyLabels(id: string[], options: { addLabels: string[]; removeLabels: string[] }) {
      await Promise.all(
        id.map(async (messageId) => {
          const message = await client
            .api(`/me/messages/${messageId}`)
            .get();

          const updatedLabels = [
            ...(message.categories || []).filter((label: string) => !options.removeLabels.includes(label)),
            ...options.addLabels,
          ];

          await client
            .api(`/me/messages/${messageId}`)
            .update({ categories: updatedLabels });
        })
      );
    },
    getDraft: async (draftId: string) => {
      try {
        const response = await client
          .api(`/me/messages/${draftId}`)
          .get();

        return response;
      } catch (error) {
        console.error("Error loading draft:", error);
        throw error;
      }
    },
    listDrafts: async (q?: string, maxResults = 20, pageToken?: string) => {
      const response = await client
        .api("/me/mailFolders/drafts/messages")
        .filter(q)
        .top(maxResults)
        .skipToken(pageToken)
        .get();

      const drafts = response.value.map((draft: any) => ({
        id: draft.id,
        threadId: draft.conversationId,
        ...parse(draft),
      }));

      return { ...response, drafts } as any;
    },
    createDraft: async (data: any) => {
      const response = await client
        .api("/me/messages")
        .post(data);

      return response;
    },
  };

  return manager;
};
