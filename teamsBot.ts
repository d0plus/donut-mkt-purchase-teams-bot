import { ActivityTypes } from "@microsoft/agents-activity";
import { BlobServiceClient } from "@azure/storage-blob";
import {
  AgentApplication,
  AttachmentDownloader,
  MemoryStorage,
  TurnContext,
  TurnState,
  CloudAdapter,
} from "@microsoft/agents-hosting";
import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import * as dotenv from "dotenv";
import { adapter } from "./index";
dotenv.config();

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_AD_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID || ""}`,
    clientSecret: process.env.AZURE_AD_CLIENT_SECRET || "",
  },
};
const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING || "";
const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER || "";
const version = "0.2.14";

interface ConversationState {
  count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

const downloader = new AttachmentDownloader();
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
  fileDownloaders: [downloader],
});

// --- Utility: streamToString ---
export async function streamToString(readableStream: NodeJS.ReadableStream): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: any[] = [];
    readableStream.on("data", (chunk) => {
      chunks.push(chunk);
    });
    readableStream.on("end", () => {
      resolve(Buffer.concat(chunks).toString("utf-8"));
    });
    readableStream.on("error", reject);
  });
}

// --- Blob Insert Helper ---
async function insertToBlob(talkerInfo: any) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient("teamsTalkerData.json");

  let existingData: any[] = [];
  try {
    const downloadBlockBlobResponse = await blockBlobClient.download();
    const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
    existingData = JSON.parse(downloaded);
    if (!Array.isArray(existingData)) existingData = [];
  } catch (e) {}

  const getUserKey = (item: any) =>
    (item.from?.id || "") + "|" + (item.conversation?.id || "");
  const userKey = getUserKey(talkerInfo);

  const userMap = new Map<string, any>();
  for (const item of existingData) {
    userMap.set(getUserKey(item), item);
  }
  userMap.set(userKey, talkerInfo);

  const dedupedData = Array.from(userMap.values());
  const content = JSON.stringify(dedupedData, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
}

// --- 10s Echo Interval Map ---

// --- Core Bot Handlers ---

// /reset: clear conversation state
teamsBot.message("/reset", async (context: TurnContext, state: ApplicationTurnState) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

// /count: show message count
teamsBot.message("/count", async (context: TurnContext, state: ApplicationTurnState) => {
  const count = state.conversation.count ?? 0;
  await context.sendActivity(`The count is ${count}`);
});

// /diag: show activity
teamsBot.message("/diag", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(context.activity));
});

// /state: show state
teamsBot.message("/state", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(state));
});

// /runtime: show runtime info
teamsBot.message("/runtime", async (context: TurnContext, state: ApplicationTurnState) => {
  const runtime = {
    nodeversion: process.version,
    sdkversion: version,
  };
  await context.sendActivity(JSON.stringify(runtime));
});

// Welcome on membersAdded
teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity(
      `Hi there! I'm an echo bot running on Agents SDK version ${version} that will echo what you said to me.`
    );
  }
);

// Email type: auto reply
teamsBot.activity(
  (context) => context.activity.type === "email" ||
    (context.activity.channelData && context.activity.channelData.email),
  async (context, state) => {
    await context.sendActivity("i got it");
  }
);

// Main message handler: reply, insert to blob, and start 10s echo
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

    // 若 activity.channelId === "msteams" 且 from.role === "user"，代表 Teams 使用者主動發訊息，僅回覆固定文字
    // 僅當 activity.type === "message" 時才回覆
    if (context.activity.type === "message") {
      if (context.activity.channelId === "msteams" && context.activity.from?.role === "user") {
        await context.sendActivity("我收到你的訊息，目前為您服務中");
        return;
      } else {
        // 僅 web post 或其他來源才 echo 原文
        await context.sendActivity(context.activity.text || "我收到你的訊息，目前為您服務中");
      }
    }

    // 只有非 Teams 使用者主動訊息才執行 blob 更新
    const reference = {
      serviceUrl: context.activity.serviceUrl,
      channelId: context.activity.channelId,
      conversation: {
        id: context.activity.conversation?.id,
        tenantId: context.activity.conversation?.tenantId ?? context.activity.channelData?.tenant?.id,
        conversationType: context.activity.conversation?.conversationType
      },
      bot: context.activity.recipient,
      user: context.activity.from
    };
    const talkerInfo = {
      time: new Date().toISOString(),
      reference,
      from: context.activity.from,
      conversation: context.activity.conversation,
      channelId: context.activity.channelId,
      serviceUrl: context.activity.serviceUrl,
      recipient: context.activity.recipient,
      teamsChannelId: context.activity.channelData?.channel?.id || null,
      teamsTeamId: context.activity.channelData?.team?.id || null,
      conversationType: context.activity.conversation?.conversationType || null,
      tenantId: context.activity.conversation?.tenantId || null,
      id: context.activity.id,
      replyToId: context.activity.replyToId,
      summary: {
        userId: context.activity.from?.id,
        aadObjectId: context.activity.from?.aadObjectId,
        conversationId: context.activity.conversation?.id,
        teamsTeamId: context.activity.channelData?.team?.id || null,
        teamsChannelId: context.activity.channelData?.channel?.id || null,
      }
    };
    await insertToBlob(talkerInfo);

    // Prepare conversation reference for proactive send
    const userKey = (context.activity.from?.id || "") + "|" + (context.activity.conversation?.id || "");
    // 已移除 10 秒 echo interval，這裡不再有任何 interval 處理
  }
);

// Regex and function-based fallback handlers
teamsBot.activity(/^message/, async (context: TurnContext, state: ApplicationTurnState) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});
teamsBot.activity(
  async (context: TurnContext) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);
