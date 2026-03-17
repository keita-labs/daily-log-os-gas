/**
 * スプレッドシートを開いたときに実行される処理
 * カスタムメニューを作成する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('OS Doc化')
    .addItem('ドキュメント作成実行', 'executeManual')
    .addItem('設定チェック', 'checkConfigManual')
    .addSeparator()
    .addItem('トリガー設定案内', 'showTriggerGuide')
    .addToUi();
}

/**
 * 手動実行用のエントリポイント
 */
function executeManual() {
  executeProcess(false);
}

/**
 * トリガー（定期実行）用のエントリポイント
 */
function executeTrigger() {
  executeProcess(true);
}

/**
 * 設定チェックを手動実行する機能
 */
function checkConfigManual() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = loadConfig();
    validateConfig(config);
    validateSheets(config);
    ui.alert('設定チェック完了', '設定およびシートの構成に問題はありません。処理を実行可能です。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('設定エラー', '以下の問題が見つかりました:\n\n' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * トリガー設定の案内を表示する
 */
function showTriggerGuide() {
  const ui = SpreadsheetApp.getUi();
  const guideText = '定期実行を行うには、以下の手順でトリガーを設定してください。\n\n' +
                    '1. 画面上部のメニュー「拡張機能」>「Apps Script」を開く\n' +
                    '2. 左側の時計アイコン（トリガー）をクリック\n' +
                    '3. 右下の「トリガーを追加」をクリック\n' +
                    '4. 実行する関数: 「executeTrigger」を選択\n' +
                    '5. イベントのソース: 「時間主導型」を選択\n' +
                    '6. タイプ: 「日付ベースのタイマー」を選択\n' +
                    '7. configシートの trigger_hour で指定した時間帯を選択して保存';
  ui.alert('設定手順案内', guideText, ui.ButtonSet.OK);
}

/**
 * 実際のメイン処理プロセス
 * @param {boolean} isTrigger - トリガー実行かどうかのフラグ
 */
function executeProcess(isTrigger) {
  let ui = null;
  if (!isTrigger) {
    try {
      ui = SpreadsheetApp.getUi();
    } catch (e) {
      // Webからの実行などUIがないコンテキストの場合は無視
    }
  }
  
  try {
    // 1. 設定読み込みと検証
    const config = loadConfig();
    validateConfig(config);
    validateSheets(config);
    
    // トリガー実行時、機能が無効化されていればスキップ
    if (isTrigger && config.trigger_enabled.toUpperCase() !== 'TRUE') {
      console.log('定期実行機能がオフ(trigger_enabled != TRUE)のため、処理を終了します。');
      return;
    }
    
    // 2. 本処理実行
    const summary = processDiarySheet(config);
    
    // 3. 結果表示
    const msg = `処理が完了しました。\n\n` + 
                `対象件数: ${summary.total}件\n` + 
                `成功　　: ${summary.success}件\n` + 
                `スキップ: ${summary.skip}件\n` + 
                `エラー　: ${summary.fail}件`;
    
    if (!isTrigger && ui) {
      ui.alert('実行完了', msg, ui.ButtonSet.OK);
    }
    console.log(msg);

  } catch (e) {
    const errorMsg = '処理中に致命的なエラーが発生し、停止しました:\n\n' + e.message;
    if (!isTrigger && ui) {
      ui.alert('エラー停止', errorMsg, ui.ButtonSet.OK);
    }
    console.error(errorMsg);
  }
}
