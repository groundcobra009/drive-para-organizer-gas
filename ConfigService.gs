/**
 * ConfigService
 * 設定情報の保存・読み込みを担当するサービス
 * PropertiesServiceを使用してスクリプトプロパティに保存
 */

var ConfigService = {
  // プロパティキー
  KEYS: {
    PROJECT_FOLDER_ID: 'PROJECT_FOLDER_ID',
    AREA_FOLDER_ID: 'AREA_FOLDER_ID',
    RESOURCE_FOLDER_ID: 'RESOURCE_FOLDER_ID',
    ARCHIVE_FOLDER_ID: 'ARCHIVE_FOLDER_ID'
  },
  
  /**
   * 設定を取得
   * @return {Object} 設定オブジェクト
   */
  getConfig: function() {
    var properties = PropertiesService.getScriptProperties();

    return {
      projectFolderId: properties.getProperty(this.KEYS.PROJECT_FOLDER_ID) || '',
      areaFolderId: properties.getProperty(this.KEYS.AREA_FOLDER_ID) || '',
      resourceFolderId: properties.getProperty(this.KEYS.RESOURCE_FOLDER_ID) || '',
      archiveFolderId: properties.getProperty(this.KEYS.ARCHIVE_FOLDER_ID) || ''
    };
  },
  
  /**
   * 設定を保存
   * @param {Object} config - 設定オブジェクト
   */
  saveConfig: function(config) {
    try {
      var properties = PropertiesService.getScriptProperties();

      if (config.projectFolderId !== undefined) {
        properties.setProperty(this.KEYS.PROJECT_FOLDER_ID, config.projectFolderId || '');
      }
      if (config.areaFolderId !== undefined) {
        properties.setProperty(this.KEYS.AREA_FOLDER_ID, config.areaFolderId || '');
      }
      if (config.resourceFolderId !== undefined) {
        properties.setProperty(this.KEYS.RESOURCE_FOLDER_ID, config.resourceFolderId || '');
      }
      if (config.archiveFolderId !== undefined) {
        properties.setProperty(this.KEYS.ARCHIVE_FOLDER_ID, config.archiveFolderId || '');
      }

      Logger.log('設定を保存しました');
      return { success: true, message: '設定を保存しました' };
    } catch (error) {
      Logger.log('saveConfig エラー: ' + error.toString());
      return { success: false, message: '設定の保存に失敗しました: ' + error.toString() };
    }
  },
  
  /**
   * 設定をクリア
   */
  clearConfig: function() {
    try {
      var properties = PropertiesService.getScriptProperties();
      properties.deleteAllProperties();
      
      Logger.log('設定をクリアしました');
      return { success: true, message: '設定をクリアしました' };
    } catch (error) {
      Logger.log('clearConfig エラー: ' + error.toString());
      return { success: false, message: '設定のクリアに失敗しました: ' + error.toString() };
    }
  }
};
