// checkAmount.ts
import axios from "axios";

/**
 * @param data 可選，傳送的資料內容
 */
export async function triggerCheckAmount(data?: any): Promise<any> {
  const endpoint = "https://d4e116f2e9c3.ngrok-free.app/check-amount";
  if (!endpoint) {
    throw new Error("CHECK_AMOUNT_ENDPOINT 環境變數未設定");
  }
  try {
    return await axios.post(endpoint, data ?? {});
  } catch (err) {
    console.error("[checkAmount] POST 失敗", err);
    throw err;
  }
}