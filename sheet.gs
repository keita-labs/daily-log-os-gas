/**
 * シート構成と必須ヘッダの存在チェックを行う
 * @param {Object} config 
 */
function validateSheets(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ターゲットシート検証
  const diarySheet = ss.getSheetByName(config.target_sheet_name);
  if (!diarySheet) throw new Error(`対象の「${config.target_sheet_name}」シートが見つかりません。`);
  
  // ログシート検証
  const logSheet = ss.getSheetByName(config.log_sheet_name);
  if (!logSheet) throw new Error(`ログ用の「${config.log_sheet_name}」シートが見つかりません。`);
  
  // ターゲットシートヘッダ検証
  const headers = diarySheet.getRange(1, 1, 1, Math.max(diarySheet.getLastColumn(), 1)).getValues()[0];
  const requiredHeaders = [
    config.header_date, config.header_thought, config.header_action, config.header_knowledge,
    config.header_doc_status, config.header_thought_url, config.header_action_url, config.header_knowledge_url
  ];
  
  const missingHeaders = requiredHeaders.filter(h => headers.indexOf(h) === -1);
  if (missingHeaders.length > 0) {
    throw new Error(`「${config.target_sheet_name}」シートに以下の必須ヘッダが存在しません:\n` + missingHeaders.join(', '));
  }

  // ログシート簡易ヘッダ検証
  const logHeaders = logSheet.getRange(1, 1, 1, Math.max(logSheet.getLastColumn(), 1)).getValues()[0];
  const logRequired = ['実行日時', '結果', '詳細'];
  const logMissing = logRequired.filter(h => logHeaders.indexOf(h) === -1);
  if (logMissing.length > 0) {
    throw new Error(`「${config.log_sheet_name}」シートに必要なヘッダ（実行日時、結果、詳細 など）が不足しています。`);
  }
}

/**
 * 対象シートを読み込み、条件に合致する行のドキュメントを作成・更新する
 * @param {Object} config 
 * @return {Object} 処理結果サマリオブジェクト
 */
function processDiarySheet(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.target_sheet_name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 処理効率化のため列インデックスを取得しておく
  const colIdx = {
    date: headers.indexOf(config.header_date),
    thought: headers.indexOf(config.header_thought),
    action: headers.indexOf(config.header_action),
    knowledge: headers.indexOf(config.header_knowledge),
    status: headers.indexOf(config.header_doc_status),
    thoughtUrl: headers.indexOf(config.header_thought_url),
    actionUrl: headers.indexOf(config.header_action_url),
    knowledgeUrl: headers.indexOf(config.header_knowledge_url)
  };
  
  let summary = { total: 0, success: 0, skip: 0, fail: 0 };
  
  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1; // スプレッドシート側の行番号(1始まり)
    const row = data[i];
    const docStatus = row[colIdx.status];
    
    // 対象ステータスの場合のみ処理
    if (docStatus === config.status_target) {
      summary.total++;
      
      const rawDate = row[colIdx.date];
      const valThought = String(row[colIdx.thought] || '').trim();
      const valAction = String(row[colIdx.action] || '').trim();
      const valKnowledge = String(row[colIdx.knowledge] || '').trim();
      
      let urlThought = String(row[colIdx.thoughtUrl] || '').trim();
      let urlAction = String(row[colIdx.actionUrl] || '').trim();
      let urlKnowledge = String(row[colIdx.knowledgeUrl] || '').trim();
      
      // 1. 必須空欄スキップチェック
      if (!valThought || !valAction || !valKnowledge) {
        writeLog(config, rowNum, rawDate, 'SKIP', '思考/行動/知識のいずれかが空欄のため未作成', urlThought, urlAction, urlKnowledge);
        summary.skip++;
        continue; // 処理をスキップして次の行へ
      }
      
      // 2. 日付パース
      const dateStr = formatDateForFileName(rawDate);
      if (!dateStr) {
        writeLog(config, rowNum, rawDate, 'FAIL', '日付データのパースに失敗(不正なフォーマット)', urlThought, urlAction, urlKnowledge);
        summary.fail++;
        continue; 
      }

      // 3. ドキュメント作成処理（行ごとにtry-catchして他行に影響を与えない）
      try {
        let isUpdated = false;
        
        // 思考OS
        if (!urlThought) {
          urlThought = createDocInFolder(config.thought_folder_id, dateStr, valThought);
          sheet.getRange(rowNum, colIdx.thoughtUrl + 1).setValue(urlThought);
          isUpdated = true;
        }
        
        // 行動OS
        if (!urlAction) {
          urlAction = createDocInFolder(config.action_folder_id, dateStr, valAction);
          sheet.getRange(rowNum, colIdx.actionUrl + 1).setValue(urlAction);
          isUpdated = true;
        }
        
        // 知識OS
        if (!urlKnowledge) {
          urlKnowledge = createDocInFolder(config.knowledge_folder_id, dateStr, valKnowledge);
          sheet.getRange(rowNum, colIdx.knowledgeUrl + 1).setValue(urlKnowledge);
          isUpdated = true;
        }
        
        // 4. ステータス更新（すべてURLが埋まったら完了とする）
        if (urlThought && urlAction && urlKnowledge) {
          sheet.getRange(rowNum, colIdx.status + 1).setValue(config.status_done);
        }
        
        const logDetail = isUpdated ? 'ドキュメントを作成しました' : 'すでに作成済み(URL存在)のためスキップ状態での完了';
        writeLog(config, rowNum, dateStr, 'SUCCESS', logDetail, urlThought, urlAction, urlKnowledge);
        summary.success++;
        
      } catch (e) {
        writeLog(config, rowNum, dateStr, 'FAIL', e.message, urlThought, urlAction, urlKnowledge);
        summary.fail++;
      }
    }
  }
  
  return summary;
}
