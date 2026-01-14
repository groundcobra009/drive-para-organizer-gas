/**
 * UI
 * サイドバーUIとの通信を担当する関数
 */

/**
 * 設定を取得（サイドバーから呼び出し）
 * @return {Object} 設定オブジェクト
 */
function getConfig() {
  return ConfigService.getConfig();
}

/**
 * 設定を保存（サイドバーから呼び出し）
 * @param {Object} config - 設定オブジェクト
 * @return {Object} 結果オブジェクト
 */
function saveConfig(config) {
  return ConfigService.saveConfig(config);
}

/**
 * フォルダIDの検証（サイドバーから呼び出し）
 * @param {string} folderId - 検証するフォルダID
 * @return {Object} 検証結果
 */
function validateFolderId(folderId) {
  if (!folderId || folderId.trim() === '') {
    return { valid: false, message: 'フォルダIDが入力されていません' };
  }
  
  var isValid = DriveService.validateFolderId(folderId);
  
  if (isValid) {
    try {
      var folderName = DriveService.getFolderName(folderId);
      return { 
        valid: true, 
        message: '有効なフォルダです: ' + folderName 
      };
    } catch (error) {
      return { valid: true, message: '有効なフォルダです' };
    }
  } else {
    return { valid: false, message: '無効なフォルダIDです' };
  }
}
