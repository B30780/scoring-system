/**
 * 競賽評分系統 - Google Apps Script 後端
 * 功能：接收評分網頁送來的資料，寫入 Google 試算表。
 *
 * 部署步驟請參考同資料夾的 README.md
 */

// 試算表分頁名稱（若你的分頁不叫「評分紀錄」，請改成實際名稱）
const SHEET_NAME = "評分紀錄";

/**
 * 接收 POST 請求並寫入試算表
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    sheet.appendRow([
      new Date(),                 // 伺服器接收時間
      data.judge || "",           // 評審姓名
      data.contestant || "",      // 參賽者
      data.creativity,            // 創意性
      data.technical,             // 技術性
      data.bug,                   // BUG 尋找力
      totalScore_(data),          // 總分
      data.comment || "",         // 意見
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * 瀏覽器直接開啟網址時顯示狀態（方便測試部署是否成功）
 */
function doGet() {
  return json_({ ok: true, message: "評分系統後端運作中" });
}

/**
 * 取得目標分頁，若不存在則自動建立並加上標題列
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "接收時間", "評審", "參賽者",
      "創意性", "技術性", "BUG尋找力",
      "總分", "意見",
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function totalScore_(d) {
  return (Number(d.creativity) || 0) + (Number(d.technical) || 0) + (Number(d.bug) || 0);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
