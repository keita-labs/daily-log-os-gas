/**
 * ログシートへ1行分の結果を追記する
 */
function writeLog(config, rowNum, rawDateStr, result, detail, thoughtUrl, actionUrl, knowledgeUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(config.log_sheet_name);
  if (!logSheet) return; // 事前チェックを通っているので基本的に実行される
  
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  
  // logシートの列(A〜I)の想定順番に合わせる
  const logRow = [
    timestamp,                 // A: 実行日時
    config.target_sheet_name,  // B: 対象シート
    rowNum,                    // C: 行番号
    rawDateStr || '',          // D: 日付
    result,                    // E: 結果
    detail,                    // F: 詳細
    thoughtUrl || '',          // G: 思考DocURL
    actionUrl || '',           // H: 行動DocURL
    knowledgeUrl || ''         // I: 知識DocURL
  ];
  
  logSheet.appendRow(logRow);
}
