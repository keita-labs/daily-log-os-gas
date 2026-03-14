/**
 * 指定したフォルダにドキュメントを作成し、URLを返す
 * @param {string} folderId 
 * @param {string} dateStr YYYYMMDD形式（そのままファイル名になります）
 * @param {string} content 書き込む本文
 * @return {string} 作成されたドキュメントのURL
 */
function createDocInFolder(folderId, dateStr, content) {
  let targetFolder;
  try {
    targetFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error(`フォルダIDへのアクセス失敗(${folderId}): 存在しないか権限がありません`);
  }
  
  // バージョン検索せず、日付そのものをファイル名として新規作成
  const doc = DocumentApp.create(dateStr);
  doc.getBody().setText(content);
  doc.saveAndClose();
  
  // 作成したファイルを対象フォルダへ移動
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(targetFolder);
  
  return doc.getUrl();
}
