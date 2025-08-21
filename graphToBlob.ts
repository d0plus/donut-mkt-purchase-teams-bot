/* Teams account to Azure Blob Storage */
import { BlobServiceClient } from "@azure/storage-blob";
import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import * as dotenv from "dotenv";
dotenv.config();


/* Azure AD config */
const msalConfig = {
  auth: {
    clientId: process.env.AZURE_AD_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID!}`,
    clientSecret: process.env.AZURE_AD_CLIENT_SECRET!,
  },
};

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
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const content = JSON.stringify(data, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content));
  console.log(`Uploaded ${blobName} to ${process.env.AZURE_BLOB_CONTAINER}`);
}

async function main() {
  const token = await getGraphToken();
  if (!token) throw new Error("Cannot get Graph API token");
  // Read blob and send batch messages
  await sendMessageToAllUsersInBlob(token, "teamsUser.json");
}

// CLI test entry
if (require.main === module) {
  // main().catch(console.error);
  sendTestMessageToAllUsersInTokensBlob("tokens.json", "This is a proactive test message from the bot.").catch(console.error);
}
import "isomorphic-fetch";

/* Proactive send test message to all users in tokens.json */
export async function sendTestMessageToAllUsersInTokensBlob(blobName: string = "tokens.json", message: string = "This is a proactive test message from the bot.") {
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const downloadBlockBlobResponse = await blockBlobClient.download(0);
  const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
  const tokens = JSON.parse(downloaded); // { [userId]: { accessToken, ... } }
  for (const userId of Object.keys(tokens)) {
    const entry = tokens[userId];
    if (entry.accessToken && entry.userInfo?.id) {
      try {
        await sendTeamsMessageWithUserToken(entry.accessToken, entry.userInfo.id, message);
        console.log(`Sent proactive message to userId=${entry.userInfo.id}`);
      } catch (e) {
        console.error(`Send message failed userId=${entry.userInfo.id}`, e);
      }
    }
  }
}

/* Proactive send Teams message (application-only token, require existing conversationId) */
export async function sendTeamsMessageToConversation(appAccessToken: string, conversationId: string, message: string) {
  const client = Client.init({
    authProvider: (done) => {
      done(null, appAccessToken);
    },
  });

  await client.api(`/chats/${conversationId}/messages`).post({
    body: {
      content: message
    }
  });
}
/* Read all users from Azure Blob and send message via Graph API */
export async function sendMessageToAllUsersInBlob(appAccessToken: string, blobName: string) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  const downloadBlockBlobResponse = await blockBlobClient.download(0);
  const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
  const users = JSON.parse(downloaded); // [{ conversationId: "xxx" }, ...]
  for (const user of users) {
    if (user.conversationId) {
      await sendTeamsMessageToConversation(appAccessToken, user.conversationId, "I am calling the graph API");
    }
  }
}

/* Utility: stream to string */
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
/* Send Teams message with user-delegated access token */
export async function sendTeamsMessageWithUserToken(accessToken: string, userId: string, message: string) {
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });

  // Get 1:1 chatId with user
  const chats = await client.api(`/users/${userId}/chats`).get();
  let chatId: string | undefined = undefined;
  if (chats.value && chats.value.length > 0) {
    chatId = chats.value.find((c: any) => c.chatType === "oneOnOne")?.id;
  }
  if (!chatId) {
    // Create new chat if not exist
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

  // Send message
  await client.api(`/chats/${chatId}/messages`).post({
    body: {
      content: message
    }
  });
}