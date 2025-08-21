import { BlobServiceClient } from "@azure/storage-blob";

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
const TEST_USER_ID = "29:n tQ4-I_mzDn-Bc1P1F-g4REhQnsW0wUvjshl6duLE4v78twa4RsopLh8jyad-p0ReV_g-VinYKJci5x9se6jQm4A";
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
import { sendTeamsMessageWithUserToken } from "./graphToBlob";

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
  * 取得 POST 資料，從 blob 比對 staffEmail，發送 Teams 訊息
  * 主流程：根據 postData.staffEmail 發送 Teams 訊息
  */
export async function handlePostAndNotifyStaff(postData: any) {
  // 顯示所有收到的欄位與型別
  console.log("[handlePostAndNotifyStaff] 收到來自網站的 POST 資料:", postData);
  if (postData && typeof postData === "object") {
    Object.keys(postData).forEach((k) => {
      console.log(`[handlePostAndNotifyStaff] 欄位: ${k}, 型別: ${typeof postData[k]}, 值:`, postData[k]);
    });
  } else {
    console.log("[handlePostAndNotifyStaff] postData 不是物件，實際型別:", typeof postData);
  }
  // 依據欄位內容組合訊息
  const staffEmail = postData.staffEmail || "";
  const content = postData.content;
  const text = postData.text;
  const message = typeof text === "string" && text.trim() ? text : (typeof content === "string" && content.trim() ? content : "you got order");
  try {
    const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING!;
    const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER!;
    const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
    const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
    const blockBlobClient = containerClient.getBlockBlobClient("tokens.json");
    const downloadBlockBlobResponse = await blockBlobClient.download(0);
    const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
    const tokens = JSON.parse(downloaded); // { [userId]: { accessToken, userInfo: { email, id } } }
    let found = false;
    for (const userId of Object.keys(tokens)) {
      const entry = tokens[userId];
      if (entry.userInfo?.email?.toLowerCase() === staffEmail.toLowerCase()) {
        console.log(`[handlePostAndNotifyStaff] 找到對應 userId: ${userId}, email: ${entry.userInfo.email}`);
        await sendTeamsMessageWithUserToken(entry.accessToken, entry.userInfo.id, message);
        console.log(`已發送訊息給 ${staffEmail} (userId=${entry.userInfo.id})，訊息內容: ${message}`);
        found = true;
        break;
      }
    }
    if (!found) {
      console.error("[handlePostAndNotifyStaff] 找不到對應的 staffEmail:", staffEmail);
    }
  } catch (err) {
    console.error("[handlePostAndNotifyStaff] 發生例外錯誤:", err);
  }
}
