/**
 * 現場報名接收端 — Google Apps Script Web App
 *
 * 部署步驟：
 * 1. 開啟本活動的 Google Sheets
 * 2. 選單 → 擴充功能 → Apps Script
 * 3. 貼上此程式碼（覆蓋原有內容）
 * 4. 存檔後點選「部署」→「新增部署作業」
 *    - 類型：Web 應用程式
 *    - 執行身分：我（服務帳號擁有者）
 *    - 誰可以存取：所有人
 * 5. 點「部署」→ 複製「網頁應用程式網址」
 * 6. 將此網址填入 Streamlit 的 .streamlit/secrets.toml：
 *    [gas]
 *    registration_url = "https://script.google.com/macros/s/...../exec"
 */

var SHEET_NAME = "表單回應 1";
var HEADERS    = ["時間戳記", "姓名", "聯繫方式", "年齡", "地址", "有無求道", "得知管道"];

function doPost(e) {
  // 使用 ScriptLock 確保並發安全（同時多人送出時排隊處理）
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);   // 等待最多 15 秒

  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // 若工作表不存在則自動建立並加表頭
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
           .setFontWeight("bold")
           .setBackground("#4CAF50")
           .setFontColor("white");
    }

    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" }),
      data["姓名"]    || "",
      data["聯繫方式"] || "",
      data["年齡"]    || "",
      data["地址"]    || "",
      data["有無求道"] || "無",
      data["得知管道"] || "其他",
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

// 用於測試 Web App 是否正常運作（瀏覽器直接開啟 URL 時呼叫）
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", sheet: SHEET_NAME }))
    .setMimeType(ContentService.MimeType.JSON);
}
