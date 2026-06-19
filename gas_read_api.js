/**
 * 義診系統 — 資料讀取 API
 *
 * 安裝位置：義診系統資料庫試算表（1yYWTlQ6qTOz-7R9w8tMo292oAvRP6yi8oFSugXbzXYQ）
 *   擴充功能 → Apps Script → 貼上此程式碼 → 儲存 → 部署為 Web 應用程式
 *   執行身分：我（擁有者）
 *   存取權限：所有人（含匿名）
 *
 * 呼叫方式：
 *   GET {url}?sheet=Settings
 *   GET {url}?sheet=Queue
 *   GET {url}?sheet=Registration
 */

var ALLOWED_SHEETS = ["Settings", "Registration", "Queue", "Tasks", "Roles", "Equipment"];

function doGet(e) {
  try {
    var sheetName = e.parameter.sheet;
    if (!sheetName) {
      return _json({ error: "缺少 sheet 參數" }, 400);
    }
    if (ALLOWED_SHEETS.indexOf(sheetName) < 0) {
      return _json({ error: "不允許讀取：" + sheetName }, 403);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName(sheetName);
    if (!ws) {
      return _json({ error: "找不到工作表：" + sheetName }, 404);
    }

    var range = ws.getDataRange();
    var values = range.getValues();

    if (values.length === 0) {
      return _json({ headers: [], rows: [] });
    }

    var headers = values[0].map(function(h) { return String(h); });
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var val = values[i][j];
        // Date 物件轉 ISO 字串
        if (val instanceof Date) {
          row[headers[j]] = Utilities.formatDate(val, "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
        } else {
          row[headers[j]] = val === null || val === undefined ? "" : val;
        }
      }
      rows.push(row);
    }

    return _json({ headers: headers, rows: rows });

  } catch (err) {
    return _json({ error: err.toString() }, 500);
  }
}

function _json(obj, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
