// orderCheckByDate.ts
import { TurnContext } from "@microsoft/agents-hosting";
import axios from "axios";

function parseDateRange(input: string): { start: string, end: string } | null {
  const match = input.match(/(\d{2}-\d{2}-\d{4})\s*(?:至|to)\s*(\d{2}-\d{2}-\d{4})/i);
  if (!match) return null;
  const [ , start, end ] = match;
  // 轉換為 ISO 格式
  const [d1, m1, y1] = start.split("-");
  const [d2, m2, y2] = end.split("-");
  return {
    start: `${y1}-${m1}-${d1}T00:00:00Z`,
    end: `${y2}-${m2}-${d2}T23:59:59Z`
  };
}

export async function handleOrderCheckByDate(context: TurnContext, state: any, dateInput: string) {
  const staffEmail = state.conversation.staffEmail;
  const range = parseDateRange(dateInput);
  if (!range) {
    await context.sendActivity("日期格式錯誤，請重新輸入（格式如：31-12-2024 至 01-05-2025）");
    return;
  }
  try {
    const resp = await axios.post("https://4dd94d1be57f.ngrok-free.app/option/all", {
      staffEmail,
      startDate: range.start,
      endDate: range.end
    });
    if (resp.data && resp.data.orders && Array.isArray(resp.data.orders) && resp.data.orders.length > 0) {
      const card = {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
        version: "1.4",
        body: [
          {
            type: "TextBlock",
            text: `查詢區間：${dateInput}`,
            weight: "Bolder",
            size: "Large",
            color: "Accent",
            horizontalAlignment: "Center"
          },
          ...resp.data.orders.map((order: any, idx: number) => ({
            type: "Container",
            items: [
              {
                type: "TextBlock",
                text: `📝 訂單 #${idx + 1}`,
                weight: "Bolder",
                size: "Medium",
                color: "Good",
                spacing: "Small"
              },
              {
                type: "FactSet",
                facts: [
                  { title: "客戶", value: order.clientName },
                  { title: "PO", value: order.poNumber },
                  { title: "金額", value: order.amount },
                  { title: "建立時間", value: new Date(order.createdAt).toLocaleString("zh-TW", { timeZone: "Asia/Shanghai" }) }
                ]
              },
              {
                type: "TextBlock",
                text: "──────────────",
                color: "Accent",
                spacing: "Small",
                isSubtle: true
              }
            ]
          }))
        ]
      };
      await context.sendActivity({
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card
          }
        ]
      } as any);
    } else {
      await context.sendActivity("此區間內查無訂單資料。");
    }
  } catch (err) {
    await context.sendActivity("查詢失敗，請稍後再試。");
  }
}