/**
 * DriveService
 * Google Drive API操作を担当するサービス
 */

var DriveService = {
  /**
   * My Drive直下のファイル一覧を取得
   * @return {Array} ファイル情報の配列
   */
  getMyDriveFiles: function() {
    try {
      var rootFolder = DriveApp.getRootFolder();
      var files = rootFolder.getFiles();
      var fileList = [];

      // フォルダを除外してファイルのみを取得
      while (files.hasNext()) {
        var file = files.next();
        // MIMEタイプがフォルダでないことを確認
        if (file.getMimeType() !== 'application/vnd.google-apps.folder') {
          fileList.push({
            id: file.getId(),
            name: file.getName(),
            mimeType: file.getMimeType(),
            modifiedTime: file.getLastUpdated(),
            size: file.getSize(),
            url: file.getUrl()
          });
        }
      }

      // 更新日時でソート（新しい順）
      fileList.sort(function(a, b) {
        return b.modifiedTime.getTime() - a.modifiedTime.getTime();
      });

      return fileList;
    } catch (error) {
      Logger.log('getMyDriveFiles エラー: ' + error.toString());
      throw error;
    }
  },
  
  /**
   * ファイルを移動
   * @param {string} fileId - 移動するファイルのID
   * @param {string} targetFolderId - 移動先フォルダのID
   */
  moveFile: function(fileId, targetFolderId) {
    try {
      var file = DriveApp.getFileById(fileId);
      var targetFolder = DriveApp.getFolderById(targetFolderId);
      
      // ファイルを移動
      file.moveTo(targetFolder);
      
      Logger.log('ファイル移動成功: ' + file.getName() + ' -> ' + targetFolder.getName());
    } catch (error) {
      Logger.log('moveFile エラー: ' + error.toString());
      throw new Error('ファイル移動に失敗しました: ' + error.toString());
    }
  },
  
  /**
   * ファイル名を変更
   * @param {string} fileId - 変更するファイルのID
   * @param {string} newName - 新しいファイル名
   */
  renameFile: function(fileId, newName) {
    try {
      var file = DriveApp.getFileById(fileId);
      var oldName = file.getName();
      file.setName(newName);
      
      Logger.log('ファイル名変更成功: ' + oldName + ' -> ' + newName);
    } catch (error) {
      Logger.log('renameFile エラー: ' + error.toString());
      throw new Error('ファイル名変更に失敗しました: ' + error.toString());
    }
  },
  
  /**
   * フォルダIDの有効性を確認
   * @param {string} folderId - 確認するフォルダのID
   * @return {boolean} 有効な場合true
   */
  validateFolderId: function(folderId) {
    try {
      if (!folderId || folderId.trim() === '') {
        return false;
      }
      
      var folder = DriveApp.getFolderById(folderId.trim());
      return folder !== null;
    } catch (error) {
      Logger.log('validateFolderId エラー: ' + error.toString());
      return false;
    }
  },
  
  /**
   * フォルダ名を取得
   * @param {string} folderId - フォルダのID
   * @return {string} フォルダ名
   */
  getFolderName: function(folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      return folder.getName();
    } catch (error) {
      Logger.log('getFolderName エラー: ' + error.toString());
      return '';
    }
  }
};
