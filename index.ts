import { BlobServiceClient } from "@azure/storage-blob";

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
  // Log errors to console (consider Azure app insights in production)
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages
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

/* Utility: NodeJS readable stream to string */
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

/* POST: find staffEmail in blob and send Teams message */
export async function handlePostAndNotifyStaff(postData: any) {
  console.log("[handlePostAndNotifyStaff] Received POST data:", postData);
  if (postData && typeof postData === "object") {
    Object.keys(postData).forEach((k) => {
      console.log(`[handlePostAndNotifyStaff] Field: ${k}, Type: ${typeof postData[k]}, Value:`, postData[k]);
    });
  } else {
    console.log("[handlePostAndNotifyStaff] postData is not object, actual type:", typeof postData);
  }
  const staffEmail = postData.staffEmail || "";
  const content = postData.content;
  const text = postData.text;
  const message = typeof text === "string" && text.trim() ? text : (typeof content === "string" && content.trim() ? content : "you got order");
  try {
    const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
    const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
    const blockBlobClient = containerClient.getBlockBlobClient("tokens.json");
    const downloadBlockBlobResponse = await blockBlobClient.download(0);
    const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
    const tokens = JSON.parse(downloaded); // { [userId]: { accessToken, userInfo: { email, id } } }
    let found = false;
    for (const userId of Object.keys(tokens)) {
      const entry = tokens[userId];
      if (entry.userInfo?.email?.toLowerCase() === staffEmail.toLowerCase()) {
        console.log(`[handlePostAndNotifyStaff] Found userId: ${userId}, email: ${entry.userInfo.email}`);
        await sendTeamsMessageWithUserToken(entry.accessToken, entry.userInfo.id, message);
        console.log(`Sent message to ${staffEmail} (userId=${entry.userInfo.id}), message: ${message}`);
        found = true;
        break;
      }
    }
    if (!found) {
      console.error("[handlePostAndNotifyStaff] staffEmail not found:", staffEmail);
    }
  } catch (err) {
    console.error("[handlePostAndNotifyStaff] Exception:", err);
  }
}
