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

async function getGraphToken() {
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
  const user = await getUserByEmail(userEmail, token);
  await uploadJsonToBlob("teamsUser.json", user);
}

if (require.main === module) {
  main().catch(console.error);
}