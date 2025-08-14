// 範例：將資料上傳至 Azure Blob Storage（botsimon2blob 容器）
import { BlobServiceClient } from "@azure/storage-blob";
import * as dotenv from "dotenv";
dotenv.config();

const AZURE_STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING!;
const AZURE_BLOB_CONTAINER = process.env.AZURE_BLOB_CONTAINER!;

export async function uploadTextToBlob(blobName: string, content: string) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);
  await blockBlobClient.upload(content, Buffer.byteLength(content));
  console.log(`已上傳 ${blobName} 至 ${AZURE_BLOB_CONTAINER}`);
}
// 新增：根據 teamsid upsert 資料到 blob
/**
 * 根據 id 欄位（如 teamInfo.id）去重複，所有資料集中存於一個 blob。
 */
/**
 * 以 aadObjectId 為唯一鍵去重複，確保資料唯一且可用於後續啟動新對話。
 * 請確保 teamInfo 物件有 aadObjectId、conversation 相關欄位。
 */
/**
 * 以 aadObjectId 為唯一鍵去重複，僅儲存最新資料，移除 message 欄位，避免多筆重複。
 * 請確保 teamInfo 物件有 aadObjectId、conversation 相關欄位。
 */
/**
 * 以 from.id（Teams user id）為唯一鍵，每位 user 只會有一筆紀錄，避免重複。
 * 請確保 teamInfo 物件有 from.id 欄位。
 */
/**
 * 強制只保留每位 user 一筆資料（from.id 為唯一鍵），徹底去除所有重複。
 */
export async function upsertTeamInfoToBlob(blobName: string, teamInfo: { from: { id: string }, [key: string]: any }) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
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

  // 以 from.id + tenantId 為唯一鍵，確保同一 Teams 使用者跨租戶也不重複
  // 移除 msg/message 欄位，且不儲存用戶輸入內容
  const sanitizedInfo = { ...teamInfo };
  delete sanitizedInfo.msg;
  delete sanitizedInfo.message;
  // 只保留必要欄位
  // 你可根據實際需求調整保留欄位
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
  // 不存 msg/message
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

  // 建立 Map 以 userKey 唯一
  const userMap = new Map<string, any>();
  for (const item of data) {
    userMap.set(getUserKey(item), item);
  }
  userMap.set(userKey, infoToSave); // 覆蓋或新增

  const deduped = Array.from(userMap.values());

  const content = JSON.stringify(deduped, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
  console.log(`已 upsert ${userKey} 至 ${blobName}`);

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

// 測試用：將 "hello world" 上傳為 test.txt
if (require.main === module) {
  uploadTextToBlob("test.txt", "hello world").catch(console.error);
}
/**
 * 工具函式：將 blob 內容從多個 JSON 物件修正為 JSON 陣列格式，僅保留每 user 一筆。
 */
export async function fixBlobToArrayFormat(blobName: string) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_BLOB_CONTAINER);
  const blockBlobClient = containerClient.getBlockBlobClient(blobName);

  let raw = "";
  try {
    const downloadBlockBlobResponse = await blockBlobClient.download();
    raw = await streamToString(downloadBlockBlobResponse.readableStreamBody);
  } catch (e) {
    console.error("無法讀取 blob，請確認 blob 是否存在");
    return;
  }

  // 嘗試將多個 JSON 物件合併為陣列
  let arr: any[] = [];
  try {
    arr = JSON.parse(raw);
    if (!Array.isArray(arr)) throw new Error();
  } catch {
    // 若不是陣列，嘗試以正則分割並 parse
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

  // 以 from.id + tenantId 唯一去重
  const getUserKey = (item: any) =>
    (item.from?.id || "") + "|" + (item.conversation?.tenantId || "");
  const userMap = new Map<string, any>();
  for (const item of arr) {
    userMap.set(getUserKey(item), item);
  }
  const deduped = Array.from(userMap.values());

  // 覆蓋回 blob
  const content = JSON.stringify(deduped, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
  console.log(`已修正 ${blobName} 為陣列格式並去重`);
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