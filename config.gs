/**
 * configシートから設定値を読み込みオブジェクトとして返す
 * @return {Object} config設定オブジェクト
 */
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('config');
  if (!sheet) {
    throw new Error('「config」シートが見つかりません。作成してください。');
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  // 1行目はヘッダ想定（key, value, 備考）なので i=1から開始
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    const value = String(data[i][1]).trim();
    if (key) {
      config[key] = value;
    }
  }
  return config;
}

/**
 * configオブジェクトの必須キーと値の形式を検証する
 * @param {Object} config 
 */
function validateConfig(config) {
  // 必須キー一覧
  const requiredKeys = [
    'target_sheet_name', 'thought_folder_id', 'action_folder_id', 'knowledge_folder_id',
    'status_target', 'status_done', 'trigger_enabled', 'trigger_hour', 'log_sheet_name',
    'header_date', 'header_thought', 'header_action', 'header_knowledge',
    'header_doc_status', 'header_thought_url', 'header_action_url', 'header_knowledge_url'
  ];
  
  const missingKeys = requiredKeys.filter(key => !(key in config) || config[key] === '');
  if (missingKeys.length > 0) {
    throw new Error('configシートに以下の必須設定が見つからないか値が空です:\n' + missingKeys.join(', '));
  }
  
  // triggers_hour チェック
  const hour = parseInt(config.trigger_hour, 10);
  if (isNaN(hour) || hour < 0 || hour > 23) {
    throw new Error('config設定エラー：trigger_hour は 0 から 23 の数値を指定してください。');
  }

  // フォルダID簡易チェック（DriveAppでのアクセス可能性はdrive側で検証する前提）
  if (config.thought_folder_id.length < 15 || config.action_folder_id.length < 15 || config.knowledge_folder_id.length < 15) {
     throw new Error('config設定エラー：一部のフォルダIDの形式が不正です（短すぎます）。');
  }
}
