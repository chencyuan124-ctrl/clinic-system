// ==========================================
// 義診掛號系統 — 場次重置工具
// 貼到 Google Sheet 的 Apps Script 編輯器
// ==========================================

// 建立自訂選單
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ 系統管理")
    .addItem("📋 查看各工作表筆數", "showSummary")
    .addSeparator()
    .addItem("🗑️ 清除排隊名單 (Queue)", "clearQueue")
    .addItem("🔄 重置報名人數 & 清空老師名單 (Settings)", "resetSettings")
    .addItem("🗑️ 清除報名歷史 (Registration)", "clearRegistration")
    .addSeparator()
    .addItem("🔥 一鍵全部重置（新場次）", "resetAll")
    .addToUi();
}

// ── 查看摘要 ──────────────────────────────
function showSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ["Queue", "Registration", "Settings"];
  var msg = "📊 目前各工作表資料筆數：\n\n";

  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      var rows = sheet.getLastRow() - 1; // 扣除標題列
      msg += "• " + name + "：" + Math.max(rows, 0) + " 筆\n";
    } else {
      msg += "• " + name + "：（找不到工作表）\n";
    }
  });

  SpreadsheetApp.getUi().alert(msg);
}

// ── 清除排隊名單 ──────────────────────────
function clearQueue() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "⚠️ 確認清除",
    "即將刪除所有排隊叫號名單（Queue）。\n\n此操作無法復原，確定要繼續？",
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Queue");
  if (!sheet) { ui.alert("找不到 Queue 工作表！"); return; }

  _clearSheetKeepHeader(sheet);
  ui.alert("✅ 完成！Queue 已清除。");
}

// ── 重置設定（歸零報名數、清空老師名單）──
function resetSettings() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "⚠️ 確認重置",
    "即將將 Settings 的「已報名數」全部歸零，並清空「老師名單」。\n\n確定要繼續？",
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if (!sheet) { ui.alert("找不到 Settings 工作表！"); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { ui.alert("Settings 目前無資料。"); return; }

  // 找出欄位位置
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var teacherCol = headers.indexOf("老師名單") + 1;
  var countCol   = headers.indexOf("已報名數") + 1;

  if (teacherCol > 0) sheet.getRange(2, teacherCol, lastRow - 1, 1).clearContent();
  if (countCol   > 0) sheet.getRange(2, countCol,   lastRow - 1, 1).setValue(0);

  ui.alert("✅ 完成！Settings 已重置。");
}

// ── 清除報名歷史 ──────────────────────────
function clearRegistration() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "⚠️ 確認清除",
    "即將刪除所有民眾報名歷史紀錄（Registration）。\n\n⚠️ 請先至系統下載備份！\n\n確定要繼續？",
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration");
  if (!sheet) { ui.alert("找不到 Registration 工作表！"); return; }

  _clearSheetKeepHeader(sheet);
  ui.alert("✅ 完成！Registration 已清除。");
}

// ── 一鍵全部重置 ──────────────────────────
function resetAll() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "🔥 全場次重置",
    "即將執行以下三項操作：\n\n" +
    "1. 清除所有排隊名單（Queue）\n" +
    "2. 重置報名人數與老師名單（Settings）\n" +
    "3. 清除所有報名歷史（Registration）\n\n" +
    "⚠️ 請確認已在系統內下載備份！\n\n" +
    "確定要執行全部重置？",
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var qSheet = ss.getSheetByName("Queue");
  if (qSheet) _clearSheetKeepHeader(qSheet);

  var sSheet = ss.getSheetByName("Settings");
  if (sSheet) {
    var lastRow = sSheet.getLastRow();
    if (lastRow > 1) {
      var headers = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getValues()[0];
      var teacherCol = headers.indexOf("老師名單") + 1;
      var countCol   = headers.indexOf("已報名數") + 1;
      if (teacherCol > 0) sSheet.getRange(2, teacherCol, lastRow - 1, 1).clearContent();
      if (countCol   > 0) sSheet.getRange(2, countCol,   lastRow - 1, 1).setValue(0);
    }
  }

  var rSheet = ss.getSheetByName("Registration");
  if (rSheet) _clearSheetKeepHeader(rSheet);

  ui.alert("✨ 全部重置完成！系統已準備好迎接下一場活動。");
}

// ── 工具函式：清除資料列但保留標題 ─────────
function _clearSheetKeepHeader(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}
