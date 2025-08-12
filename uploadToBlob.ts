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

// 測試用：將 "hello world" 上傳為 test.txt
if (require.main === module) {
  uploadTextToBlob("test.txt", "hello world").catch(console.error);
}