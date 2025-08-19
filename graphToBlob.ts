// 取得 Teams 帳號資料並寫入 Azure Blob Storage
import { BlobServiceClient } from "@azure/storage-blob";
import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import * as dotenv from "dotenv";
dotenv.config();

const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING!;
const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER!;

// 請將下列資訊填入您的 Azure AD 註冊應用程式資訊
const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID",
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID",
    clientSecret: "YOUR_CLIENT_SECRET",
  },
};
const userEmail = "simon@something.com"; // 可改為動態取得

export async function getGraphToken() {
  const cca = new ConfidentialClientApplication(msalConfig);
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return result?.accessToken;
}

async function getUserByEmail(email: string, accessToken: string) {
  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });
  const user = await client.api(`/users/${email}`).get();
  return user;
}

async function uploadJsonToBlob(blobName: string, data: any) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const content = JSON.stringify(data, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content));
  console.log(`已上傳 ${blobName} 至 ${AZURE_BLOB_CONTAINER}`);
}

async function main() {
  const token = await getGraphToken();
  if (!token) throw new Error("無法取得 Graph API 權杖");
  // 讀取 blob 並批次推播訊息
  await sendMessageToAllUsersInBlob(token, "teamsUser.json");
}

if (require.main === module) {
  // main().catch(console.error);
  sendTestMessageToAllUsersInTokensBlob("tokens.json", "This is a proactive test message from the bot.").catch(console.error);
}
// 使用 Microsoft Graph API 主動發送 Teams 訊息給指定 userId
import "isomorphic-fetch";

// 主動發送測試訊息給所有 tokens.json 內的 user
export async function sendTestMessageToAllUsersInTokensBlob(blobName: string = "tokens.json", message: string = "This is a proactive test message from the bot.") {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const downloadBlockBlobResponse = await blockBlobClient.download(0);
  const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
  const tokens = JSON.parse(downloaded); // { [userId]: { accessToken, ... } }
  for (const userId of Object.keys(tokens)) {
    const entry = tokens[userId];
    if (entry.accessToken && entry.userInfo?.id) {
      try {
        await sendTeamsMessageWithUserToken(entry.accessToken, entry.userInfo.id, message);
        console.log(`已主動發送訊息給 userId=${entry.userInfo.id}`);
      } catch (e) {
        console.error(`發送訊息給 userId=${entry.userInfo.id} 失敗`, e);
      }
    }
  }
}

/**
 * 主動發送 Teams 訊息（改為 application-only token，僅支援已存在的 conversationId）
 * @param accessToken Bot 應用程式權杖
 * @param conversationId 目標 Teams 使用者的 conversationId
 * @param message 要發送的訊息內容
 */
export async function sendTeamsMessageToConversation(appAccessToken: string, conversationId: string, message: string) {
  const client = Client.init({
    authProvider: (done) => {
      done(null, appAccessToken);
    },
  });

  // 僅能對已存在的 conversationId 發送訊息
  await client.api(`/chats/${conversationId}/messages`).post({
    body: {
      content: message
    }
  });
}
/**
 * 從 Azure Blob 讀取所有 user，並用 Graph API 發送訊息
 */
export async function sendMessageToAllUsersInBlob(appAccessToken: string, blobName: string) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const downloadBlockBlobResponse = await blockBlobClient.download(0);
  const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
  const users = JSON.parse(downloaded); // 假設格式為 [{ conversationId: "xxx" }, ...]
  for (const user of users) {
    if (user.conversationId) {
      await sendTeamsMessageToConversation(appAccessToken, user.conversationId, "I am calling the graph API");
    }
  }
}

// 工具函式：stream 轉 string
async function streamToString(readableStream: NodeJS.ReadableStream | null): Promise<string> {
  if (!readableStream) return "";
  return new Promise((resolve, reject) => {
    const chunks: any[] = [];
    readableStream.on("data", (data) => {
      chunks.push(data.toString());
    });
    readableStream.on("end", () => {
      resolve(chunks.join(""));
    });
    readableStream.on("error", reject);
  });
}
/**
 * 用 user-delegated access token 發送 Teams 訊息給指定 userId
 */

/**
 * @param accessToken 使用者授權取得的 access token
 * @param userId 目標 Teams 使用者的 userId
 * @param message 要發送的訊息內容
 */
export async function sendTeamsMessageWithUserToken(accessToken: string, userId: string, message: string) {
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  // 取得與 user 的 1:1 chatId
  const chats = await client.api(`/users/${userId}/chats`).get();
  let chatId: string | undefined = undefined;
  if (chats.value && chats.value.length > 0) {
    chatId = chats.value.find((c: any) => c.chatType === "oneOnOne")?.id;
  }
  if (!chatId) {
    // 若沒有現有 chat，則建立一個
    const chat = await client.api("/chats").post({
      chatType: "oneOnOne",
      members: [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
        }
      ]
    });
    chatId = chat.id;
  }

  // 發送訊息
  await client.api(`/chats/${chatId}/messages`).post({
    body: {
      content: message
    }
  });
}