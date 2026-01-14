/**
 * drive-para-organizer-gas
 * My Drive直下のファイルをスプレッドシートに展開し、
 * PARA分類フラグとリネーム指示に基づいて一括移動・名前変更を行う整理ツール
 */

/**
 * スプレッドシート起動時にメニューを追加
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Drive整理')
    .addItem('ファイル一覧を取得', 'loadFilesFromDrive')
    .addItem('設定を開く', 'showSidebar')
    .addItem('チェック済みファイルを移動', 'moveCheckedFiles')
    .addSeparator()
    .addItem('フォルダID列を表示/非表示', 'toggleFileIdColumn')
    .addSeparator()
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * メイン処理：My Drive直下のファイルをスプレッドシートに展開
 */
function loadFilesFromDrive() {
  try {
    var sheet = getOrCreateSheet();
    var files = DriveService.getMyDriveFiles();
    
    // ヘッダーを設定
    var headers = ['移動', 'ファイル名', 'ファイルID', '種類', '更新日時', 'サイズ', '移動先', 'URL'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    
    // データを書き込み
    if (files.length > 0) {
      var data = files.map(function(file) {
        return [
          false, // チェックボックス（移動フラグ）
          file.name,
          file.id,
          file.mimeType,
          Utilities.formatDate(file.modifiedTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
          file.size || '',
          '', // 移動先（空欄）
          file.url
        ];
      });

      sheet.getRange(2, 1, data.length, headers.length).setValues(data);

      // チェックボックス列を設定
      var checkboxRange = sheet.getRange(2, 1, data.length, 1);
      checkboxRange.insertCheckboxes();

      // 移動先列（G列=7列目）にデータ検証（ドロップダウン）を設定
      var destinationRange = sheet.getRange(2, 7, data.length, 1);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['プロジェクト', 'エリア', 'リソース', 'アーカイブ'], true)
        .setAllowInvalid(false)
        .build();
      destinationRange.setDataValidation(rule);
      
      SpreadsheetApp.getUi().alert('ファイル一覧を取得しました。\n件数: ' + files.length + '件');
    } else {
      SpreadsheetApp.getUi().alert('My Drive直下にファイルが見つかりませんでした。');
    }
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * チェック済みファイルを移動
 */
function moveCheckedFiles() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var config = ConfigService.getConfig();
    
    if (!config.projectFolderId && !config.areaFolderId &&
        !config.resourceFolderId && !config.archiveFolderId) {
      SpreadsheetApp.getUi().alert('設定が完了していません。\n「設定を開く」からフォルダIDを設定してください。');
      return;
    }
    
    var movedCount = 0;
    var errorCount = 0;
    
    // ヘッダー行をスキップして処理
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var isChecked = row[0]; // 移動フラグ
      var fileName = row[1];
      var fileId = row[2]; // C列：ファイルID
      var destination = row[6]; // G列：移動先

      if (isChecked === true && fileId && destination) {
        var targetFolderId = getTargetFolderId(config, destination);

        if (targetFolderId) {
          try {
            DriveService.moveFile(fileId, targetFolderId);
            movedCount++;
            // チェックを外す
            sheet.getRange(i + 1, 1).setValue(false);
          } catch (error) {
            Logger.log('移動エラー: ' + fileName + ' - ' + error.toString());
            errorCount++;
          }
        } else {
          Logger.log('移動先フォルダが設定されていません: ' + destination);
          errorCount++;
        }
      }
    }
    
    var message = '移動が完了しました。\n移動: ' + movedCount + '件';
    if (errorCount > 0) {
      message += '\nエラー: ' + errorCount + '件';
    }
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}

/**
 * 移動先フォルダIDを取得
 * @param {Object} config - 設定オブジェクト
 * @param {string} destination - 移動先 (プロジェクト/エリア/リソース/アーカイブ)
 * @return {string} フォルダID
 */
function getTargetFolderId(config, destination) {
  if (!destination) return null;

  switch (destination) {
    case 'プロジェクト':
      return config.projectFolderId;
    case 'エリア':
      return config.areaFolderId;
    case 'リソース':
      return config.resourceFolderId;
    case 'アーカイブ':
      return config.archiveFolderId;
    default:
      return null;
  }
}

/**
 * セル編集時のトリガー関数
 * B列（ファイル名）が編集された場合、自動的にファイル名を変更
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();

    // 「ファイル一覧」シート以外は処理しない
    if (sheet.getName() !== 'ファイル一覧') {
      return;
    }

    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();

    // B列（ファイル名）が編集された場合
    if (col === 2 && row > 1) {
      var newFileName = range.getValue();
      var fileId = sheet.getRange(row, 3).getValue(); // C列：ファイルID

      if (fileId && newFileName && newFileName.trim() !== '') {
        try {
          DriveService.renameFile(fileId, newFileName.trim());
          Logger.log('ファイル名を変更しました: ' + newFileName);
        } catch (error) {
          Logger.log('リネームエラー: ' + fileId + ' - ' + error.toString());
          SpreadsheetApp.getUi().alert('ファイル名の変更に失敗しました: ' + error.toString());
          // エラー時は元の値に戻す
          e.range.setValue(e.oldValue);
        }
      }
    }
  } catch (error) {
    Logger.log('onEdit エラー: ' + error.toString());
  }
}

/**
 * サイドバーを表示
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Drive整理 - 設定')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * ヘルプを表示
 */
function showHelp() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Drive整理 - ヘルプ')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * シートを取得または作成
 */
function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ファイル一覧');

  if (!sheet) {
    sheet = ss.insertSheet('ファイル一覧');
  }

  return sheet;
}

/**
 * フォルダID列（D列）の表示/非表示を切り替え
 */
function toggleFileIdColumn() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var column = 4; // D列 = ファイルID列

    // 現在の表示状態を確認
    var isHidden = sheet.isColumnHiddenByUser(column);

    if (isHidden) {
      // 非表示→表示
      sheet.showColumns(column);
      SpreadsheetApp.getUi().alert('フォルダID列を表示しました。');
    } else {
      // 表示→非表示
      sheet.hideColumns(column);
      SpreadsheetApp.getUi().alert('フォルダID列を非表示にしました。');
    }
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.toString());
  }
}
