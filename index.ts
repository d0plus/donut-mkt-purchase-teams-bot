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
const adapter = new CloudAdapter(authConfig);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  if (context.activity.type === "message") {
    await context.sendTraceActivity(
      "OnTurnError Trace",
      `${error}`,
      "https://www.botframework.com/schemas/error",
      "TurnError"
    );
    await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
};

adapter.onTurnError = onTurnErrorHandler;

const server = express();
server.use(express.json());

import { uploadTextToBlob } from "./UploadToBlob";
import { v4 as uuidv4 } from "uuid";
import { sendTeamsMessageWithUserToken } from "./graphToBlob";

server.use("/api/webpost", (req, res, next) => {
  const authHeader = req.headers["authorization"];
  const expected = `Bearer ${process.env.NOTIFY_API_KEY}`;
  if (!authHeader || authHeader !== expected) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }
  next();
});
server.use("/api/notify", (req, res, next) => {
  const authHeader = req.headers["authorization"];
  const expected = `Bearer ${process.env.NOTIFY_API_KEY}`;
  if (!authHeader || authHeader !== expected) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }
  next();
});

import fs from "fs/promises";

import { BlobServiceClient } from "@azure/storage-blob";

server.post("/api/webpost", async (req: Request, res: Response) => {
  const authHeader = req.headers["authorization"];
  const expected = `Bearer ${process.env.NOTIFY_API_KEY}`;
  if (!authHeader || authHeader !== expected) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }
  try {
    // 完全比照 botsimon：只通知 staff，不寫入 webpost blob
    
    const staffEmail = req.body.staffEmail || "";
    const text = req.body.text;
    const content = req.body.content;
    // webpost 直接取 text 欄位內容
    const message = typeof text === "string" ? text : "you got order";
    

    // 讀取 tokens.json from Azure blob
    const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
    const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
    const blockBlobClient = containerClient.getBlockBlobClient("tokens.json");
    const downloadBlockBlobResponse = await blockBlobClient.download(0);
    const downloaded = await (async function streamToString(readableStream: NodeJS.ReadableStream | null): Promise<string> {
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
    })(downloadBlockBlobResponse.readableStreamBody);

    const tokens = JSON.parse(downloaded); // { [userId]: { accessToken, userInfo: { email, id } } }
    let found = false;
    for (const userId of Object.keys(tokens)) {
      const entry = tokens[userId];
      if (
        typeof entry.userInfo?.email === "string" &&
        entry.userInfo.email.toLowerCase() === String(staffEmail).toLowerCase()
      ) {
        console.log("[/api/webpost] 發送 Teams message 詳細資訊：", {
          userId,
          accessToken: entry.accessToken,
          teamsUserId: entry.userInfo.id,
          message
        });
        await sendTeamsMessageWithUserToken(entry.accessToken, entry.userInfo.id, message);
        found = true;
        break;
      }
    }
    if (found) {
      return res.status(200).json({ ok: true, message: "已通知 staff", staffEmail });
    } else {
      return res.status(404).json({ ok: false, error: "找不到對應 staffEmail", staffEmail });
    }
  } catch (err) {
    return res.status(500).json({ ok: false, error: (err as Error).message });
  }
});

// 新增 /api/notify 路由，含 Authorization 驗證
server.post("/api/notify", async (req: Request, res: Response) => {
  const authHeader = req.headers["authorization"];
  const expected = `Bearer ${process.env.NOTIFY_API_KEY}`;
  if (!authHeader || authHeader !== expected) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }
  // 這裡可加入 Teams Bot 處理邏輯
  res.status(200).json({ ok: true, message: "Notify received" });
});

server.post("/api/messages", async (req: Request, res: Response) => {
  await adapter.process(req, res, async (context) => {
    await teamsBot.run(context);
  });
});

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