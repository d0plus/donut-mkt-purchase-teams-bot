import { ActivityTypes } from "@microsoft/agents-activity";
import {
  AgentApplication,
  AttachmentDownloader,
  MemoryStorage,
  TurnContext,
  TurnState,
} from "@microsoft/agents-hosting";``
import { version } from "@microsoft/agents-hosting/package.json";

interface ConversationState {
  count: number;
}
type ApplicationTurnState = TurnState<ConversationState>;

const downloader = new AttachmentDownloader();

// Define storage and application
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
  fileDownloaders: [downloader],
});

// Listen for user to say '/reset' and then delete conversation state
teamsBot.message("/reset", async (context: TurnContext, state: ApplicationTurnState) => {
  state.deleteConversationState();
  await context.sendActivity("Ok I've deleted the current conversation state.");
});

teamsBot.message("/count", async (context: TurnContext, state: ApplicationTurnState) => {
  const count = state.conversation.count ?? 0;
  await context.sendActivity(`The count is ${count}`);
});

teamsBot.message("/diag", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(context.activity));
});

teamsBot.message("/state", async (context: TurnContext, state: ApplicationTurnState) => {
  await state.load(context, storage);
  await context.sendActivity(JSON.stringify(state));
});

teamsBot.message("/runtime", async (context: TurnContext, state: ApplicationTurnState) => {
  const runtime = {
    nodeversion: process.version,
    sdkversion: version,
  };
  await context.sendActivity(JSON.stringify(runtime));
});

teamsBot.conversationUpdate(
  "membersAdded",
  async (context: TurnContext, state: ApplicationTurnState) => {
    await context.sendActivity(
      `Hi there! I'm an echo bot running on Agents SDK version ${version} that will echo what you said to me.`
    );
  }
);

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
teamsBot.activity(
  ActivityTypes.Message,
  async (context: TurnContext, state: ApplicationTurnState) => {
    // Increment count state
    let count = state.conversation.count ?? 0;
    state.conversation.count = ++count;

    // Log received Teams message
    console.log("[teamsBot] 收到 Teams 訊息：", {
      from: context.activity.from,
      text: context.activity.text,
      conversation: context.activity.conversation,
      channelId: context.activity.channelId,
      timestamp: context.activity.timestamp
    });

    // 若來自 webPost，直接回覆 payload.text
    if (context.activity.channelData && context.activity.channelData.webPost && context.activity.text) {
      await context.sendActivity(context.activity.text);
    } else {
      // 回覆固定繁體中文訊息
      await context.sendActivity("我已收到你的訊息，目前機器人運作正常");
    }

    // 新增：收到 Teams 訊息時寫入 blob
    try {
      const { from, text, conversation, channelId, timestamp } = context.activity;
      // 直接靜態 import，避免動態 import 路徑錯誤
      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const { uploadTextToBlob } = require("./UploadToBlob");
      const blobName = `teamsmsg-${Date.now()}.json`;
      // 取得 email 欄位（如 botsimon 專案）
      let email = "";
      if (from?.aadObjectId) {
        try {
          // 這裡可根據實際情境呼叫 Graph API 查詢 email，或直接從 activity 取得
          // Teams activity 沒有 email 欄位，這裡可擴充查詢
          email = ""; // 若有查詢 email 的邏輯可補上
        } catch {}
      }
      // 依 botsimon 專案補齊完整欄位
      // 依 botsimon upsertTeamInfoToBlob 邏輯，每 user 只存一份且去重
      const { upsertTeamInfoToBlob } = require("./UploadToBlob");
      // 取得 email 欄位（如 botsimon 專案）
      let emailValue = "";
      if (from?.aadObjectId) {
        try {
          // 完全比照 botsimon: 只用 Graph API 查詢 email
          try {
            const { ConfidentialClientApplication } = require("@azure/msal-node");
            const { Client } = require("@microsoft/microsoft-graph-client");
            const msalConfig = {
              auth: {
                clientId: process.env.AZURE_AD_CLIENT_ID || "",
                authority: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID || ""}`,
                clientSecret: process.env.AZURE_AD_CLIENT_SECRET || "",
              },
            };
            const msal = new ConfidentialClientApplication(msalConfig);
            const result = await msal.acquireTokenByClientCredential({
              scopes: ["https://graph.microsoft.com/.default"],
            });
            const accessToken = result?.accessToken;
            if (accessToken) {
              const client = Client.init({
                authProvider: (done: any) => done(null, accessToken),
              });
              const userObj = await client.api(`/users/${from.aadObjectId}`).get();
              console.log("[teamsBot] Graph API userObj:", userObj);
              console.log("[teamsBot] userObj.mail:", userObj.mail, "userObj.userPrincipalName:", userObj.userPrincipalName);
              if (userObj.mail && userObj.mail.includes("@")) emailValue = userObj.mail;
              else if (userObj.userPrincipalName && userObj.userPrincipalName.includes("@")) emailValue = userObj.userPrincipalName;
              else emailValue = "";
            }
          } catch (e) {
            console.error("[teamsBot] 取得 email 失敗", e);
          }
        } catch {}
      }
      const teamInfo = {
        from: {
          id: from?.id || "unknown",
          name: from?.name || "",
          aadObjectId: from?.aadObjectId || "",
          email: emailValue || ""
        },
        conversation: {
          id: conversation?.id || "",
          tenantId: conversation?.tenantId || "",
          conversationType: conversation?.conversationType || ""
        },
        channelId: context.activity.channelId,
        serviceUrl: context.activity.serviceUrl,
        recipient: context.activity.recipient,
        time: context.activity.timestamp
      };
      await upsertTeamInfoToBlob("teamsUser.json", teamInfo);
      console.log("[teamsBot] 已 upsert user info to blob", teamInfo.from.id);
    } catch (err) {
      console.error("[teamsBot] blob 寫入失敗", err);
    }
  }
);

teamsBot.activity(/^message/, async (context: TurnContext, state: ApplicationTurnState) => {
  await context.sendActivity(`Matched with regex: ${context.activity.type}`);
});

teamsBot.activity(
  async (context: TurnContext) => Promise.resolve(context.activity.type === "message"),
  async (context, state) => {
    await context.sendActivity(`Matched function: ${context.activity.type}`);
  }
);
