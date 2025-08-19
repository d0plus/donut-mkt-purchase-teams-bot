/**
 * 測試主動發送訊息到所有 blob 中的 Teams 使用者
 * 需先確保 teamsTalkerData.json 已有資料
 */
import { BlobServiceClient } from "@azure/storage-blob";
import { adapter } from "./index";
import { CloudAdapter } from "@microsoft/agents-hosting";

const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING || "";
const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER || "";

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

async function sendProactiveMessages(message: string = "proactive test") {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient("teamsTalkerData.json");

  try {
    const downloadBlockBlobResponse = await blockBlobClient.download();
    const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
    const userData = JSON.parse(downloaded);

    if (!Array.isArray(userData)) {
      console.error("[Proactive] Blob data is not in the expected format.");
      return;
    }

    for (const user of userData) {
      if (user.conversation?.id && user.serviceUrl && user.recipient?.id && user.from?.id) {
        const reference = {
          serviceUrl: user.serviceUrl,
          channelId: user.channelId || "msteams",
          conversation: {
            id: user.conversation.id,
            tenantId: user.conversation.tenantId || user.tenantId,
            conversationType: user.conversation.conversationType
          },
          bot: user.recipient,
          user: user.from
        };
        try {
          await (adapter as CloudAdapter).continueConversation(reference, async (context) => {
            await context.sendActivity(`[proactive test] ${message}`);
          });
          console.log(`[Proactive] 已發送訊息給 conversationId=${user.conversation.id}`);
        } catch (error) {
          console.error(`[Proactive] 發送訊息失敗 conversationId=${user.conversation.id}`, error);
        }
      } else {
        console.warn("[Proactive] User data is missing reference fields:", user);
      }
    }
    console.log("[Proactive] 主動發訊息流程完成");
  } catch (error) {
    console.error("[Proactive] Failed to retrieve or process blob data:", error);
  }
}

// CLI 執行
if (require.main === module) {
  (async () => {
    await sendProactiveMessages("這是 CLI 主動發送測試訊息");
    process.exit(0);
  })();
}
