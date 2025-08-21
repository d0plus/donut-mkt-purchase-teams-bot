import { ActivityTypes } from "@microsoft/agents-activity";
import { BlobServiceClient } from "@azure/storage-blob";
import {
  AgentApplication,
  AttachmentDownloader,
  MemoryStorage,
  TurnContext,
  TurnState,
  CloudAdapter,
} from "@microsoft/agents-hosting";
import { Client } from "@microsoft/microsoft-graph-client";
import { ConfidentialClientApplication } from "@azure/msal-node";
import * as dotenv from "dotenv";
import { adapter } from "./index";
dotenv.config();

const msalConfig = {
  auth: {
    clientId: process.env.AZURE_AD_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID || ""}`,
    clientSecret: process.env.AZURE_AD_CLIENT_SECRET || "",
  },
};
const version = "0.2.14";

interface ConversationState {
  count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

const downloader = new AttachmentDownloader();
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
  fileDownloaders: [downloader],
});

/* Utility: stream to string */
export async function streamToString(readableStream: NodeJS.ReadableStream): Promise<string> {
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
// Get email by aadObjectId
async function getEmailByAadObjectId(aadObjectId: string): Promise<string | null> {
  const msal = new ConfidentialClientApplication(msalConfig);
  const result = await msal.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  const accessToken = result?.accessToken;
  if (!accessToken) return null;

  const client = Client.init({
    authProvider: (done) => done(null, accessToken),
  });
  try {
    const user = await client.api(`/users/${aadObjectId}`).get();
    if (user.mail && user.mail.includes("@")) return user.mail;
    if (user.userPrincipalName && user.userPrincipalName.includes("@")) return user.userPrincipalName;
    return null;
  } catch (e) {
    return null;
  }
}

// Insert to blob helper
async function insertToBlob(talkerInfo: any) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(process.env.AZURE_STORAGE_CONNECTION_STRING || "");
  const containerClient = blobServiceClient.getContainerClient(process.env.AZURE_BLOB_CONTAINER || "");

  const userId = talkerInfo.user.id || "";
  const email = talkerInfo.user.email || "";
  const fileKey = userId ? `staff_${userId}.json` : (email ? `staff_${email}.json` : `staff_unknown_${Date.now()}.json`);
  talkerInfo.rowName = userId || email;

  const blockBlobClient = containerClient.getBlockBlobClient(fileKey);
  const content = JSON.stringify(talkerInfo, null, 2);
  await blockBlobClient.upload(content, Buffer.byteLength(content), undefined);
}



// /reset: clear conversation state
teamsBot.message("/reset", async (context: TurnContext, state: ApplicationTurnState) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

// /count: show message count
teamsBot.message("/count", async (context: TurnContext, state: ApplicationTurnState) => {
  const count = state.conversation.count ?? 0;
  await context.sendActivity(`The count is ${count}`);
});

// /diag: show activity
teamsBot.message("/diag", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(context.activity));
});

// /state: show state
teamsBot.message("/state", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(state));
});

// /runtime: show runtime info
teamsBot.message("/runtime", async (context: TurnContext, state: ApplicationTurnState) => {
  const runtime = {
    nodeversion: process.version,
    sdkversion: version,
  };
  await context.sendActivity(JSON.stringify(runtime));
});

// Welcome on membersAdded
teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity(
      `Hi there! I'm an echo bot running on Agents SDK version ${version} that will echo what you said to me.`
    );
  }
);

// Email type: auto reply
teamsBot.activity(
  (context) => context.activity.type === "email" ||
    (context.activity.channelData && context.activity.channelData.email),
  async (context, state) => {
    await context.sendActivity("i got it");
  }
);

// Main message handler: reply, insert to blob
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

    if (context.activity.type === "message") {
      if (context.activity.channelData && context.activity.channelData.webPost === true) {
        await context.sendActivity(context.activity.text || "我收到你的訊息，目前運行中");
        return;
      }
      if (context.activity.channelId === "msteams" && context.activity.from?.role === "user") {
        await context.sendActivity("我收到你的訊息，目前運行中");
        return;
      }
      await context.sendActivity("我收到你的訊息，目前運行中");
    }

    const reference = {
      serviceUrl: context.activity.serviceUrl,
      channelId: context.activity.channelId,
      conversation: {
        id: context.activity.conversation?.id,
        tenantId: context.activity.conversation?.tenantId ?? context.activity.channelData?.tenant?.id,
        conversationType: context.activity.conversation?.conversationType
      },
      bot: context.activity.recipient,
      user: context.activity.from
    };
    const talkerInfo = {
      time: new Date().toISOString(),
      user: {
        id: context.activity.from?.id,
        name: context.activity.from?.name,
        aadObjectId: context.activity.from?.aadObjectId,
        email: null
      },
      conversation: {
        id: context.activity.conversation?.id,
        tenantId: context.activity.conversation?.tenantId ?? context.activity.channelData?.tenant?.id,
        type: context.activity.conversation?.conversationType || null
      },
      bot: {
        id: context.activity.recipient?.id,
        name: context.activity.recipient?.name
      },
      serviceUrl: context.activity.serviceUrl
    };
    let email: string | null = null;
    if (context.activity.from?.aadObjectId) {
      email = await getEmailByAadObjectId(context.activity.from.aadObjectId);
    }
    if (!email) {
      console.error("No email found, not writing to blob, user:", context.activity.from);
      return;
    }
    talkerInfo.user.email = email;
    await insertToBlob(talkerInfo);
  }
);

// Regex and function-based fallback handlers
teamsBot.activity(/^message/, async (context: TurnContext, state: ApplicationTurnState) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});
teamsBot.activity(
  async (context: TurnContext) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);
