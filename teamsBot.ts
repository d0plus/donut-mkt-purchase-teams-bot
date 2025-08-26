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
  staffEmail?: string;
  checkStep?: string;
}
type ApplicationTurnState = TurnState<ConversationState>;
// 包裝 state.save，遇到 eTag 衝突自動重試
async function saveStateWithRetry(state: any, context: any, storage: any, maxRetry = 3) {
  let retry = 0;
  while (retry < maxRetry) {
    try {
      await state.save(context, storage);
      return;
    } catch (err: any) {
      if (err.message && err.message.includes("eTag conflict")) {
        // 僅記錄，不傳送給使用者
        console.warn("eTag conflict, retrying...");
        await state.load(context, storage);
        retry++;
      } else {
        throw err;
      }
    }
  }
  throw new Error("state.save failed after retry due to eTag conflict");
}

const downloader = new AttachmentDownloader();

// Define storage and application
const storage = new MemoryStorage();
export const teamsBot = new AgentApplication<ApplicationTurnState>({
  storage,
  fileDownloaders: [downloader],
});

import { triggerCheckAmount } from "./checkAmount";

// Listen for user to say '/check' and trigger local POST
import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";

teamsBot.message("/check", async (context: TurnContext, state: ApplicationTurnState) => {
  // 查詢 Teams 使用者 email（Graph API by aadObjectId）
  let staffEmail = "";
  const aadObjectId = context.activity.from?.aadObjectId;
  if (aadObjectId) {
    try {
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
        const userObj = await client.api(`/users/${aadObjectId}`).get();
        if (userObj.mail && userObj.mail.includes("@")) staffEmail = userObj.mail;
        else if (userObj.userPrincipalName && userObj.userPrincipalName.includes("@")) staffEmail = userObj.userPrincipalName;
      }
    } catch (e) {
      console.error("[/check] 查詢 email 失敗", e);
    }
  }
  // 若查不到 email，預設用 Teams 名稱+ID組合
  if (!staffEmail) {
    staffEmail = context.activity.from?.name
      ? `${context.activity.from.name}_${context.activity.from.id || ""}`
      : context.activity.from?.id || "unknown";
  }

  // 顯示選單
  await context.sendActivity({
    type: "message",
    text: "請選擇查詢動作，請輸入數字：\n\n1️⃣ 查詢帳單總數\n2️⃣ 依日期區間查詢\n3️⃣ 查詢最新 X 筆訂單",
  } as any);

  // 儲存 staffEmail 於 state 供後續步驟使用
  state.conversation.staffEmail = staffEmail;
  state.conversation.checkStep = "awaitOption";
});

// 處理 /check 後續互動
teamsBot.activity(ActivityTypes.Message, async (context: TurnContext, state: ApplicationTurnState) => {
  try {
    await state.load(context, storage);
    // 若為 webPost 訊息（如新訂單通知），直接 return，不進行任何固定訊息回覆
    if (context.activity.channelData && context.activity.channelData.webPost === true && context.activity.text && context.activity.text.startsWith("有新訂單：")) {
      await context.sendActivity(context.activity.text);
      return;
    }
    // 若在 /check 流程中，不進行 fallback handler 的 state.save
    // 已移除選單流程
    if (state.conversation.checkStep === "awaitOption" && state.conversation.staffEmail) {
      const option = (context.activity.text || "").trim();
      if (option === "1") {
        // 查詢所有帳單
        const axios = require("axios");
        try {
          // 查詢全部訂單
          const resp = await axios.post("https://4dd94d1be57f.ngrok-free.app/option/all", {
            staffEmail: state.conversation.staffEmail
          });
          if (resp.data && Array.isArray(resp.data.orders)) {
            await context.sendActivity(`訂單總數：${resp.data.orders.length}`);
          } else {
            await context.sendActivity("查無訂單資料。");
          }
        } catch (err) {
          await context.sendActivity("查詢失敗，請稍後再試。");
        }
        state.conversation.checkStep = undefined;
        await saveStateWithRetry(state, context, storage);
        return;
      } else if (option === "2") {
        await context.sendActivity('請輸入查詢日期區間，格式 = "01-01-2020 to 10-01-2020"');
        state.conversation.checkStep = "awaitDateRange";
        await saveStateWithRetry(state, context, storage);
        return;
      } else if (option === "3") {
        await context.sendActivity("請輸入要查詢的最新訂單數量（數字）：");
        state.conversation.checkStep = "awaitLatestOrderCount";
        await saveStateWithRetry(state, context, storage);
        return;
      } else {
        await context.sendActivity("請輸入 1、2 或 3。");
        await saveStateWithRetry(state, context, storage);
        return;
      }
    }
    if (state.conversation.checkStep === "awaitLatestOrderCount" && state.conversation.staffEmail) {
      const countInput = (context.activity.text || "").trim();
      const axios = require("axios");
      const count = Math.max(1, parseInt(countInput, 10) || 1);
      try {
        const resp = await axios.post("https://4dd94d1be57f.ngrok-free.app/option/latest-by-email", {
          staffEmail: state.conversation.staffEmail,
          count
        });
        if (resp.data && Array.isArray(resp.data.orders) && resp.data.orders.length > 0) {
          let msg = `最新 ${count} 筆訂單：\n`;
          resp.data.orders.forEach((order: any, idx: number) => {
            msg += `#${idx + 1}\n`;
            if (order.clientName) msg += `客戶：${order.clientName}\n`;
            if (order.poNumber) msg += `PO：${order.poNumber}\n`;
            if (order.amount) msg += `金額：${order.amount}\n`;
            if (order.createdAt) msg += `建立時間：${new Date(order.createdAt).toLocaleString("zh-TW", { timeZone: "Asia/Shanghai" })}\n`;
            msg += "----------------------\n";
          });
          await context.sendActivity(msg);
        } else {
          await context.sendActivity("查無訂單資料。");
        }
      } catch (err) {
        let errMsg = "查詢失敗，請稍後再試。";
        if (err && err.message) errMsg += `\n${err.message}`;
        if (err && err.response && err.response.data) errMsg += `\n${JSON.stringify(err.response.data)}`;
        await context.sendActivity(errMsg);
      }
      state.conversation.checkStep = undefined;
      await saveStateWithRetry(state, context, storage);
      return;
    }
    if (state.conversation.checkStep === "awaitDateRange" && state.conversation.staffEmail) {
      const dateInput = (context.activity.text || "").trim();
      // 解析日期格式 "dd-mm-yyyy to dd-mm-yyyy"
      const match = dateInput.match(/^(\d{2}-\d{2}-\d{4})\s*to\s*(\d{2}-\d{2}-\d{4})$/);
      if (match) {
        // 轉換 dd-mm-yyyy 為 yyyy-mm-dd
        const [d1, m1, y1] = match[1].split("-");
        const [d2, m2, y2] = match[2].split("-");
        const startDate = new Date(`${y1}-${m1}-${d1}T00:00:00Z`).toISOString();
        const endDate = new Date(`${y2}-${m2}-${d2}T23:59:59Z`).toISOString();
        const axios = require("axios");
        try {
          const resp = await axios.post("https://4dd94d1be57f.ngrok-free.app/option/all", {
            staffEmail: state.conversation.staffEmail,
            startDate,
            endDate
          });
          if (resp.data && Array.isArray(resp.data.orders)) {
            await context.sendActivity(`區間訂單數：${resp.data.orders.length}`);
          } else {
            await context.sendActivity("查無訂單資料。");
          }
        } catch (err) {
          await context.sendActivity("查詢失敗，請稍後再試。");
        }
      } else {
        await context.sendActivity("日期格式錯誤，請重新輸入（格式：01-01-2020 to 10-01-2020）");
      }
      state.conversation.checkStep = undefined;
      await saveStateWithRetry(state, context, storage);
      return;
    }
    if (state.conversation.checkStep === "awaitCount" && state.conversation.staffEmail) {
      const countInput = (context.activity.text || "").trim();
      const { handleOrderCheckByCount } = require("./orderCheckByCount");
      await handleOrderCheckByCount(context, state, countInput);
      state.conversation.checkStep = undefined;
      await saveStateWithRetry(state, context, storage);
      return;
    }
    // 非 /check 流程才回覆固定訊息
    /* await context.sendActivity("我已收到你的訊息，機器人已啟動，你可以輸入 /check 使用額外功能"); */
  } catch (err: any) {
    if (err && err.message && err.message.includes("eTag conflict")) {
      // 僅記錄，不回傳 Teams
      console.warn("eTag conflict (outer handler), ignored.");
    } else {
      throw err;
    }
  }
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

/* 歡迎訊息已移除 */

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS

