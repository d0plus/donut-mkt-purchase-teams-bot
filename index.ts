import { BlobServiceClient } from "@azure/storage-blob";
/*
// MSAL 登入與 callback 路由
import { ConfidentialClientApplication, Configuration } from "@azure/msal-node";

const msalConfig: Configuration = {
  auth: {
    clientId: process.env.AZURE_AD_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID || ""}`,
    clientSecret: process.env.AZURE_AD_CLIENT_SECRET || "",
  },
};
const REDIRECT_URI = process.env.REDIRECT_URI || "https://bot4a8ac5.azurewebsites.net/auth/callback";
const msalClient = new ConfidentialClientApplication(msalConfig);

const SCOPES = ["User.Read", "Chat.ReadWrite", "ChatMessage.Send", "offline_access"];
*/

// 啟動時自動批次推播訊息

// 啟動時主動推播訊息
/**
 * 啟動時主動推播訊息（Bot Framework 標準 proactive message）
 * 必須等 adapter 宣告後再執行
 */
function proactiveSendAll() {
  try {
    const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING!;
    const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER!;
    const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
    const blockBlobClient = containerClient.getBlockBlobClient("teamsTalkerData.json");
    blockBlobClient.download().then(async (downloadBlockBlobResponse) => {
      const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
      const userData = JSON.parse(downloaded);

      for (const user of userData) {
        if (user.conversation?.id && user.serviceUrl && user.recipient?.id && user.from?.id) {
          const reference = {
            serviceUrl: user.serviceUrl,
            conversation: { id: user.conversation.id },
            channelId: user.channelId || "msteams",
            bot: {
              id: user.recipient.id,
              name: user.recipient.name || ""
            },
            user: {
              id: user.from.id,
              name: user.from.name || "",
              aadObjectId: user.from.aadObjectId || undefined
            }
          };
          await adapter.continueConversation(reference, async (context) => {
            console.log("[Proactive] reference.user:", reference.user);
            await context.sendActivity({
              type: "message",
              text: "這是 Bot 主動推播的訊息（proactive message）",
              from: reference.bot,
              recipient: reference.user,
              conversation: reference.conversation,
              channelId: reference.channelId,
              serviceUrl: reference.serviceUrl
            } as any);
          });
        }
      }
    }).catch((e) => {
      console.error("Proactive message failed:", e);
    });
  } catch (e) {
    console.error("Proactive message failed:", e);
  }
}

// 在 adapter 宣告後呼叫

// 測試用：定時只針對特定用戶推播
const TEST_USER_ID = "29:1Q4-I_mzDn-Bc1P1F-g4REhQnsW0wUvjshl6duLE4v78twa4RsopLh8jyad-p0ReV_g-VinYKJci5x9se6jQm4A";
const TEST_CONVERSATION_ID = "a:16__v_SUT0yXZ5YEWD6vMWfgH7YIaccfDI-UuKMMclVLVNRO-_Rs9B7zQ90k1pyDd9ss1KrEXOZoCk-Oipou_eBbgWmGkHn1PZy7IJbjfJ-WsmarT-mZ6tRHDg7uwyjaH";
const TEST_MESSAGE = "這是測試用定時主動推播訊息";

import {
  AuthConfiguration,
  authorizeJWT,
  CloudAdapter,
  loadAuthConfigFromEnv,
  Request,
  TurnContext,
} from "@microsoft/agents-hosting";
import express, { Response } from "express";

import { teamsBot } from "./teamsBot";

// Create authentication configuration
const authConfig: AuthConfiguration = loadAuthConfigFromEnv();

// Create adapter
export const adapter = new CloudAdapter(authConfig);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
  if (context.activity.type === "message") {
    // Send a trace activity
    await context.sendTraceActivity(
      "OnTurnError Trace",
      `${error}`,
      "https://www.botframework.com/schemas/error",
      "TurnError"
    );

    // Send a message to the user
    await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Create express application
const server = express();
server.use(express.json());

// MSAL 登入與 callback 路由（需放在 JWT middleware 之前）
/*
server.get("/auth/login", (req, res) => {
  const authUrl = msalClient.getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
  });
  authUrl.then((url) => res.redirect(url));
});

server.get("/auth/callback", async (req, res) => {
  const code = req.query.code as string;
  if (!code) {
    res.status(400).send("No code found in callback");
    return;
  }
  try {
    const tokenResponse = await msalClient.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
    });
    // 儲存 userId 與 accessToken 到 tokens.json
    const userId = tokenResponse.account?.homeAccountId || tokenResponse.account?.username || "unknown";
    // 儲存到 Azure Blob Storage
    const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING!;
    const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER!;
    const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
    const blockBlobClient = containerClient.getBlockBlobClient("tokens.json");
    let tokens: Record<string, any> = {};
    try {
      const downloadBlockBlobResponse = await blockBlobClient.download();
      const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
      tokens = JSON.parse(downloaded);
    } catch (e) {
      // blob 不存在時忽略
    }
    // 取得更多 Teams 使用者資訊
    let userInfo = {};
    try {
      const fetch = (await import("node-fetch")).default;
      const graphRes = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${tokenResponse.accessToken}` }
      });
      if (graphRes.ok) {
        userInfo = await graphRes.json();
      }
    } catch (e) {
      userInfo = { error: "Failed to fetch user info" };
    }
    tokens[userId] = {
      accessToken: tokenResponse.accessToken,
      expiresOn: tokenResponse.expiresOn,
      userInfo
    };
    const content = JSON.stringify(tokens, null, 2);
    await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
    // 登入成功後自動導回 Teams 或顯示可關閉視窗
    res.send(`<html>
      <body>
        <h2>登入成功，access token 已儲存到 Azure Blob。</h2>
        <script>
          if (window.opener) {
            window.opener.postMessage('auth-success', '*');
            window.close();
          }
        </script>
        <p>你可以關閉此視窗並回到 Teams。</p>
      </body>
    </html>`);
  } catch (err) {
    res.status(500).send("取得 access token 失敗：" + (err as Error).message);
  }
});
*/

/**
 * 僅對自訂 API 路由啟用 JWT 驗證，避免攔截 Bot Framework 請求
 */
server.use("/api/notify", authorizeJWT(authConfig));

// Listen for incoming requests.
server.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await teamsBot.run(context);
  });
});

// Start the server
const port = process.env.PORT || 3978;
server
  .listen(port, () => {
    console.log(
      `Bot Started, listening to port ${port} for appId ${authConfig.clientId} debug ${process.env.DEBUG}`
    );
  })
  .on("error", (err) => {
    console.error(err);
    process.exit(1);
  });

import fs from "fs";
import path from "path";

/**
 * 將 NodeJS readable stream 轉為 string
 */
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
 * === 10 秒定時 HTTP POST 主動推播訊息（取代 Bot Framework 方式） ===
 * 會讀取 blob 內所有 Teams 使用者 reference，每 10 秒對每位使用者發送一則訊息到本機 /api/messages
 * 需帶 JWT 驗證，訊息內容皆為 "這是 HTTP POST 主動推播訊息"
 */
import fetch from "node-fetch";
import jwt from "jsonwebtoken";

const POST_MESSAGE_INTERVAL = 10000; // 10 秒
const POST_MESSAGE_TEXT = "這是 HTTP POST 主動推播訊息";

async function getJwtToken() {
  // 產生 JWT token，與 authorizeJWT(authConfig) 相同邏輯
  // 這裡假設 authConfig.secret 為簽章密鑰
  const payload = {
    iss: authConfig.clientId,
    sub: authConfig.clientId,
    aud: "api://bot",
    iat: Math.floor(Date.now() / 1000),
    exp: Math.floor(Date.now() / 1000) + 60 * 5,
  };
  // 使用 clientSecret 作為 JWT 簽章密鑰
  const secret = (authConfig as any).secret || (authConfig as any).clientSecret || "secret";
  return jwt.sign(payload, secret);
}

async function postProactiveMessages() {
  try {
    const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING!;
    const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER!;
    const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
    const blockBlobClient = containerClient.getBlockBlobClient("teamsTalkerData.json");
    const downloadBlockBlobResponse = await blockBlobClient.download();
    const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
    const userData = JSON.parse(downloaded);

    const jwtToken = await getJwtToken();

    for (const user of userData) {
      // recipient 應為 bot，from 應為 user
      if (
        user.conversation?.id &&
        user.recipient?.id &&
        user.from?.id
      ) {
        const activity = {
          type: "message",
          text: POST_MESSAGE_TEXT,
          from: user.from,
          recipient: user.recipient,
          conversation: user.conversation,
          channelId: user.channelId || "msteams",
          serviceUrl: user.serviceUrl || "",
        };
        // 強制 POST 到 /api/messages endpoint，不用 user.serviceUrl
        await fetch("http://localhost:" + (process.env.PORT || 3978) + "/api/messages", {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": "Bearer " + jwtToken,
          },
          body: JSON.stringify(activity),
        });
      }
    }
  } catch (e) {
    console.error("[HTTP POST Proactive] 發送失敗：", e);
  }
}

setInterval(postProactiveMessages, POST_MESSAGE_INTERVAL);

/** 
 * === 原有 Bot Framework 主動推播程式碼已註解保留 ===
 * 
 * function proactiveSendAll() { ... }
 * ...
 */
