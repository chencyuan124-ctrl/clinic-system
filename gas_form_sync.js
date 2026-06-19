/**
 * 報名表單同步器 — Google Apps Script
 *
 * 功能：
 *   每當 Google 表單有新回應（或手動執行），
 *   自動將「姓名」「手機號碼」等欄位同步寫入主資料庫的 Registration 工作表。
 *
 * 安裝位置：
 *   開啟「報名資料試算表」→ 擴充功能 → Apps Script → 貼上此程式碼
 *
 * 設定步驟：
 *   1. 確認 TARGET_SHEET_ID 填入主資料庫試算表 ID
 *   2. 確認 SOURCE_SHEET_NAME 與表單回應的工作表名稱一致
 *   3. 儲存後，點選「觸發條件」→ 新增觸發條件：
 *      函式：syncOnFormSubmit
 *      事件來源：試算表
 *      事件類型：表單提交時
 */

// ── 設定區（請依實際情況修改）────────────────────────────────
var TARGET_SHEET_ID   = "1yYWTlQ6qTOz-7R9w8tMo292oAvRP6yi8oFSugXbzXYQ";  // 主資料庫
var SOURCE_SHEET_NAME = "表單回覆 1";  // 報名資料試算表中，表單回應的工作表名稱
var TARGET_SHEET_NAME = "Registration"; // 主資料庫中的目標工作表
// ────────────────────────────────────────────────────────────

/**
 * 表單提交觸發（設定為 onFormSubmit 觸發條件）
 */
function syncOnFormSubmit(e) {
  _doSync();
}

/**
 * 手動執行 / 時間觸發（可另外設定每分鐘或每5分鐘執行）
 */
function syncManual() {
  var result = _doSync();
  Logger.log(result);
}

// ── 核心同步邏輯 ─────────────────────────────────────────────
function _doSync() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    // 讀取來源（本試算表）
    var sourceSS    = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = sourceSS.getSheetByName(SOURCE_SHEET_NAME);
    if (!sourceSheet) return "找不到工作表：" + SOURCE_SHEET_NAME;

    var sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length <= 1) return "來源無資料";

    var srcHeaders = sourceData[0].map(function(h) { return String(h); });

    // 欄位索引
    var tsIdx    = 0; // 時間戳記固定在第1欄
    var nameIdx  = _findCol(srcHeaders, ["姓名"]);
    var phoneIdx = _findCol(srcHeaders, ["手機", "聯繫", "電話"]);
    var ageIdx   = _findCol(srcHeaders, ["年齡"]);
    var addrIdx  = _findCol(srcHeaders, ["地址"]);
    var daoIdx   = _findCol(srcHeaders, ["求道"]);
    var srcIdx   = _findCol(srcHeaders, ["管道"]);

    // 讀取目標（主資料庫 Registration）
    var targetSS    = SpreadsheetApp.openById(TARGET_SHEET_ID);
    var targetSheet = targetSS.getSheetByName(TARGET_SHEET_NAME);
    if (!targetSheet) return "找不到目標工作表：" + TARGET_SHEET_NAME;

    var targetData    = targetSheet.getDataRange().getValues();
    var targetHeaders = targetData[0].map(function(h) { return String(h); });

    var serialIdx  = targetHeaders.indexOf("報到序號");
    // 找 gform_timestamp 欄：先找標題，找不到就用最後一欄
    var gformTsIdx = targetHeaders.indexOf("gform_timestamp");
    if (gformTsIdx < 0) {
      gformTsIdx = targetHeaders.length; // 尚未建立欄時，寫到末尾
    }

    // 收集已存在的 gform_timestamp（統一轉為 ISO 字串比對）
    var existingTs = {};
    for (var i = 1; i < targetData.length; i++) {
      var cell = targetData[i][gformTsIdx];
      if (cell) {
        var tsKey = (cell instanceof Date)
          ? Utilities.formatDate(cell, "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")
          : String(cell).trim();
        if (tsKey) existingTs[tsKey] = true;
      }
    }

    // 計算目前最大序號
    var maxSerial = 0;
    for (var i = 1; i < targetData.length; i++) {
      var s = parseInt(targetData[i][serialIdx]);
      if (!isNaN(s) && s > maxSerial) maxSerial = s;
    }

    // 找出待新增的列
    var newRows  = [];
    var nowStr   = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

    for (var i = 1; i < sourceData.length; i++) {
      var row = sourceData[i];
      var rawTs = row[tsIdx];
      if (!rawTs) continue;
      // 統一轉為 ISO 格式字串做去重 key
      var tsKey = (rawTs instanceof Date)
        ? Utilities.formatDate(rawTs, "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")
        : String(rawTs).trim();
      if (!tsKey || existingTs[tsKey]) continue;  // 空白或已匯入

      maxSerial++;
      newRows.push([
        maxSerial,                                          // 報到序號
        nameIdx  >= 0 ? String(row[nameIdx]).trim()  : "", // 姓名
        ageIdx   >= 0 ? row[ageIdx]                  : "", // 年齡
        phoneIdx >= 0 ? String(row[phoneIdx]).trim() + " " : "", // 聯繫方式（加空格避免 Sheets 格式化）
        addrIdx  >= 0 ? String(row[addrIdx])         : "", // 地址
        "",                                                 // 報名項目（待工作人員分配）
        daoIdx   >= 0 ? String(row[daoIdx])          : "無", // 有無求道
        srcIdx   >= 0 ? String(row[srcIdx])          : "其他", // 得知管道
        nowStr,                                             // 報名時間
        "初次接觸",                                         // 成全進度
        tsKey,                                              // gform_timestamp（ISO 格式，防重複匯入）
      ]);
    }

    if (newRows.length === 0) return "無新資料";

    // 寫入目標工作表
    var lastRow = targetSheet.getLastRow();
    targetSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length)
               .setValues(newRows);

    return "已同步 " + newRows.length + " 筆";

  } catch (err) {
    return "錯誤：" + err.toString();
  } finally {
    lock.releaseLock();
  }
}

// ── 工具函式 ──────────────────────────────────────────────────
function _findCol(headers, keywords) {
  for (var k = 0; k < keywords.length; k++) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf(keywords[k]) >= 0) return i;
    }
  }
  return -1;
}
