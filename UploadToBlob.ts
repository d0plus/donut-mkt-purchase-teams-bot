/* Example: upload data to Azure Blob Storage */
import { BlobServiceClient } from "@azure/storage-blob";
import * as dotenv from "dotenv";
dotenv.config();

import { v4 as uuidv4 } from "uuid";

export async function uploadTextToBlob(blobName: string, data: { user: string; type: string; status: string }) {
  console.log("[uploadTextToBlob] start", { blobName, data });
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);

  // 依 botsimon 格式包裝
  const json = {
    id: uuidv4(),
    name: blobName,
    timestamp: new Date().toISOString(),
    description: "這是一個寫入 Azure Blob 的範例 JSON 檔案",
    data
  };

  const content = JSON.stringify(json, null, 2);
  try {
    await blockBlobClient.upload(content, Buffer.byteLength(content));
    console.log(`[uploadTextToBlob] Uploaded ${blobName} to ${process.env.AZURE_BLOB_CONTAINER}`);
  } catch (err) {
    console.error("[uploadTextToBlob] upload failed", err);
    throw err;
  }
}
/* Upsert data to blob by teamsid, remove duplicates by id */
export async function upsertTeamInfoToBlob(blobName: string, teamInfo: { from: { id: string }, [key: string]: any }) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);

  let data: any[] = [];
  try {
    const downloadBlockBlobResponse = await blockBlobClient.download();
    const downloaded = await streamToString(downloadBlockBlobResponse.readableStreamBody);
    data = JSON.parse(downloaded);
    if (!Array.isArray(data)) data = [];
  } catch (e) {
    data = [];
  }

  // Remove duplicates by from.id + tenantId, remove msg/message fields
  const sanitizedInfo = { ...teamInfo };
  delete sanitizedInfo.msg;
  delete sanitizedInfo.message;
  // Keep only required fields
  const {
    from,
    conversation,
    channelId,
    serviceUrl,
    recipient,
    teamsChannelId,
    teamsTeamId,
    conversationType,
    tenantId,
    id,
    summary,
    time
  } = sanitizedInfo;
  const infoToSave = {
    from,
    conversation,
    channelId,
    serviceUrl,
    recipient,
    teamsChannelId,
    teamsTeamId,
    conversationType,
    tenantId,
    id,
    summary,
    time
  };

  const getUserKey = (item: any) =>
    (item.from?.id || "") + "|" + (item.conversation?.tenantId || "");

  const userKey = getUserKey(sanitizedInfo);

  // Map by userKey
  const userMap = new Map<string, any>();
  for (const item of data) {
    userMap.set(getUserKey(item), item);
  }
  userMap.set(userKey, infoToSave); // 覆蓋或新增

  const deduped = Array.from(userMap.values());

  const content = JSON.stringify(deduped, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
  console.log(`Upserted ${userKey} to ${blobName}`);

  async function streamToString(readableStream: NodeJS.ReadableStream | null): Promise<string> {
    if (!readableStream) return "";
    return new Promise((resolve, reject) => {
      const chunks: any[] = [];
      readableStream.on("data", (data) => chunks.push(data.toString()));
      readableStream.on("end", () => resolve(chunks.join("")));
      readableStream.on("error", reject);
    });
  }
}

/* CLI test: upload "hello world" as test.txt */
if (require.main === module) {
  uploadTextToBlob("test.txt", { user: "simon", type: "sample", status: "active" }).catch(console.error);
}
/* Utility: fix blob to array format, keep one per user */
export async function fixBlobToArrayFormat(blobName: string) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING!);
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER!);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);

  let raw = "";
  try {
    const downloadBlockBlobResponse = await blockBlobClient.download();
    raw = await streamToString(downloadBlockBlobResponse.readableStreamBody);
  } catch (e) {
    console.error("Cannot read blob, please check if blob exists");
    return;
  }

  // Try merge multiple JSON objects to array
  let arr: any[] = [];
  try {
    arr = JSON.parse(raw);
    if (!Array.isArray(arr)) throw new Error();
  } catch {
    // If not array, try split and parse by regex
    arr = [];
    const matches = raw.match(/{[\s\S]*?}(?=\s*{|\s*$)/g);
    if (matches) {
      for (const m of matches) {
        try {
          arr.push(JSON.parse(m));
        } catch {}
      }
    }
  }

  // Remove duplicates by from.id + tenantId
  const getUserKey = (item: any) =>
    (item.from?.id || "") + "|" + (item.conversation?.tenantId || "");
  const userMap = new Map<string, any>();
  for (const item of arr) {
    userMap.set(getUserKey(item), item);
  }
  const deduped = Array.from(userMap.values());

  // Overwrite blob
  const content = JSON.stringify(deduped, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
  console.log(`Fixed ${blobName} to array format and deduped`);
  async function streamToString(readableStream: NodeJS.ReadableStream | null): Promise<string> {
    if (!readableStream) return "";
    return new Promise((resolve, reject) => {
      const chunks: any[] = [];
      readableStream.on("data", (data) => chunks.push(data.toString()));
      readableStream.on("end", () => resolve(chunks.join("")));
      readableStream.on("error", reject);
    });
  }
}