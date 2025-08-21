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
// --- 取得 email by aadObjectId ---
async function getEmailByAadObjectId(aadObjectId: string): Promise<string | null> {
  // 取得 application access token
  const msal = new ConfidentialClientApplication(msalConfig);
  const result = await msal.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  const accessToken = result?.accessToken;
  if (!accessToken) return null;

  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });
  try {
    const user = await client.api(`/users/${aadObjectId}`).get();
    // mail 可能為空，userPrincipalName 幾乎一定有
    if (user.mail && user.mail.includes("@")) return user.mail;
    if (user.userPrincipalName && user.userPrincipalName.includes("@")) return user.userPrincipalName;
    return null;
  } catch (e) {
    return null;
  }
}

// --- Blob Insert Helper ---
async function insertToBlob(talkerInfo: any) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);

  // 以 userId 或 email 作為檔名
  const userId = talkerInfo.user.id || "";
  const email = talkerInfo.user.email || "";
  const fileKey = userId ? `staff_${userId}.json` : (email ? `staff_${email}.json` : `staff_unknown_${Date.now()}.json`);
  talkerInfo.rowName = userId || email;

  const blockBlobClient = containerClient.getBlockBlobClient(fileKey);
  const content = JSON.stringify(talkerInfo, null, 2);
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
      // Web POST（帶有 channelData.webPost: true）→ echo text
      if (context.activity.channelData && context.activity.channelData.webPost === true) {
        await context.sendActivity(context.activity.text || "收到來自 Web 的訊息");
        return;
      }
      // Teams 使用者主動訊息 → 回覆固定訊息
      if (context.activity.channelId === "msteams" && context.activity.from?.role === "user") {
        await context.sendActivity("我收到你的訊息，目前運行中");
        return;
      }
      // 其他來源
      await context.sendActivity("我收到你的訊息，目前運行中");
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
      user: {
        id: context.activity.from?.id,
        name: context.activity.from?.name,
        aadObjectId: context.activity.from?.aadObjectId,
        email: null // 稍後填入
      },
      conversation: {
        id: context.activity.conversation?.id,
        tenantId: context.activity.conversation?.tenantId ?? context.activity.channelData?.tenant?.id,
        type: context.activity.conversation?.conversationType || null
      },
      bot: {
        id: context.activity.recipient?.id,
        name: context.activity.recipient?.name
      },
      serviceUrl: context.activity.serviceUrl
    };
    // 取得 email，若查不到則不寫入 blob
    let email: string | null = null;
    if (context.activity.from?.aadObjectId) {
      email = await getEmailByAadObjectId(context.activity.from.aadObjectId);
    }
    if (!email) {
      console.error("查無 email，不寫入 blob，user:", context.activity.from);
      return;
    }
    talkerInfo.user.email = email;
    await insertToBlob(talkerInfo);

    // Prepare conversation reference for proactive send
    // const userKey = (context.activity.from?.id || "") + "|" + (context.activity.conversation?.id || "");
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
