import { ActivityTypes } from "@microsoft/agents-activity";
import { BlobServiceClient } from "@azure/storage-blob";
import {
  AgentApplication,
  AttachmentDownloader,
  MemoryStorage,
  TurnContext,
  TurnState,
} from "@microsoft/agents-hosting";
import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import * as dotenv from "dotenv";
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

async function getGraphToken() {
  const cca = new ConfidentialClientApplication(msalConfig);
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result?.accessToken;
}

async function getUserById(userId: string, accessToken: string) {
  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });
  // userId 可能為 AAD objectId，需用 filter 查詢
  let user;
  if (userId && userId.length > 30) {
    // 以 objectId 查詢
    const users = await client.api(`/users`).filter(`id eq '${userId}'`).get();
    user = users.value && users.value.length > 0 ? users.value[0] : null;
  } else {
    // 以 UPN 查詢
    user = await client.api(`/users/${userId}`).get();
  }
  return user;
}

async function uploadJsonToBlob(blobName: string, data: any) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const content = JSON.stringify(data, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content));
}
const version = "0.2.14";

interface ConversationState {
  count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

const downloader = new AttachmentDownloader();

// Define storage and application
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
  fileDownloaders: [downloader],
});

// Listen for user to say '/reset' and then delete conversation state
teamsBot.message("/reset", async (context: TurnContext, state: ApplicationTurnState) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

teamsBot.message("/count", async (context: TurnContext, state: ApplicationTurnState) => {
  const count = state.conversation.count ?? 0;
  await context.sendActivity(`The count is ${count}`);
});

teamsBot.message("/diag", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(context.activity));
});

teamsBot.message("/state", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(state));
});

teamsBot.message("/runtime", async (context: TurnContext, state: ApplicationTurnState) => {
  const runtime = {
    nodeversion: process.version,
    sdkversion: version,
  };
  await context.sendActivity(JSON.stringify(runtime));
});

teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity(
      `Hi there! I'm an echo bot running on Agents SDK version ${version} that will echo what you said to me.`
    );
    // 取得 user id
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id;
    if (userId) {
      try {
        const token = await getGraphToken();
        if (token) {
          const userInfo = await getUserById(userId, token);
          await uploadJsonToBlob(`teamsUser_${userId}.json`, userInfo);
        }
      } catch (e) {
        console.error("寫入 Teams user info 到 blob 失敗", e);
      }
    }



  }
);

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
// 當收到 email 類型訊息時，自動回覆 "i got it"
teamsBot.activity(
  (context) => context.activity.type === "email" ||
    (context.activity.channelData && context.activity.channelData.email),
  async (context, state) => {
    await context.sendActivity("i got it");
  }
);
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    // Increment count state
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

    // 每次收到訊息時，查詢 Teams user info 並寫入 blob
    const userId = context.activity.from?.aadObjectId || context.activity.from?.id;
    if (userId) {
      try {
        const token = await getGraphToken();
        if (token) {
          const userInfo = await getUserById(userId, token);
          await uploadJsonToBlob(`teamsUser_${userId}.json`, userInfo);
        }
      } catch (e) {
        console.error("訊息時寫入 Teams user info 到 blob 失敗", e);
      }
    }

    // Echo back users request
    await context.sendActivity("我收到你的訊息");
    // 收集 Teams 訊息者資訊與會話資訊（不存 msg）
    const talkerInfo = {
      time: new Date().toISOString(),
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
    // 讀取現有 blob 資料，避免重複插入
    const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
    const blockBlobClient = containerClient.getBlockBlobClient("teamsTalkerData.json");

    let existingData: any[] = [];
    try {
      const downloadBlockBlobResponse = await blockBlobClient.download();
      const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
      existingData = JSON.parse(downloaded);
      if (!Array.isArray(existingData)) existingData = [];
    } catch (e) {
      console.log("Blob 不存在或讀取失敗，將建立新 blob");
    }

    // 以 userId + conversationId 為唯一鍵去重
    const getUserKey = (item: any) =>
      (item.from?.id || "") + "|" + (item.conversation?.id || "");
    const userKey = getUserKey(talkerInfo);

    const userMap = new Map<string, any>();
    for (const item of existingData) {
      userMap.set(getUserKey(item), item);
    }
    userMap.set(userKey, talkerInfo); // 覆蓋或新增

    const dedupedData = Array.from(userMap.values());
    const content = JSON.stringify(dedupedData, null, 2);
    await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
    console.log(`已更新 blob，避免重複插入資料`);
  }
);

teamsBot.activity(/^message/, async (context: TurnContext, state: ApplicationTurnState) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});

teamsBot.activity(
  async (context: TurnContext) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);


/**
 * Utility function to convert a readable stream to a string.
 * @param readableStream - The readable stream to convert.
 * @returns A promise that resolves to the string content of the stream.
 */
async function streamToString(readableStream: NodeJS.ReadableStream): Promise<string> {
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
