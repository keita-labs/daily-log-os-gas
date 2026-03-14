/**
 * セルに入力されている日付データを YYYYMMDD 形式の文字列に変換する
 * @param {*} val 日付オブジェクト または 文字列
 * @return {string|null} 変換成功時YYYYMMDD, 失敗時null
 */
function formatDateForFileName(val) {
  if (!val) return null;
  
  // Dateオブジェクトの場合
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyyMMdd');
  }
  
  // 文字列の場合：区切り文字をなくして8桁数字かを調べる (例: 2026/02/28 -> 20260228)
  const strVal = String(val).replace(/[-/]/g, '').trim();
  if (/^\d{8}$/.test(strVal)) {
    return strVal;
  }
  
  // 念のため文字列からDate変換を試行する
  const parsedDate = new Date(val);
  if (!isNaN(parsedDate.getTime())) {
    return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyyMMdd');
  }
  
  // パース不能な不正値はnull
  return null;
}
