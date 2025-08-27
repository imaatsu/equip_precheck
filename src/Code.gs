/**
 * GAS 始業前点検システム - Google Apps Script
 * フォーム送信をトリガーとして不適合項目をissuesシートに起票する
 */

// =============================================================================
// ユーティリティ関数群
// =============================================================================

/**
 * 日付に指定日数を加算する
 * @param {Date} date - 基準日付
 * @param {number} days - 加算日数
 * @returns {Date} 加算後の日付
 */
function addDays_(date, days) {
  var result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

/**
 * 日付をyyyy-MM-dd形式の文字列に変換する
 * @param {Date} date - 変換対象の日付
 * @returns {string} yyyy-MM-dd形式の文字列
 */
function formatDate_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 8桁のUUID断片を生成する（ID用）
 * @returns {string} 8桁の英数字
 */
function generateUuid_() {
  return Utilities.getUuid().substring(0, 8);
}

/**
 * 重複防止用の一意キーを生成する
 * @param {string} timestamp - タイムスタンプ
 * @param {string} equipmentId - 設備ID
 * @returns {string} 一意キー
 */
function generateUniqueKey_(timestamp, equipmentId) {
  if (!equipmentId) {
    return null;
  }
  // タイムスタンプ欠損時のフォールバック
  var ts = timestamp || new Date().toISOString();
  return ts + "_" + equipmentId;
}

/**
 * Script PropertiesからDEFAULT_DEADLINE_DAYSを取得する（未設定時は7日）
 * @returns {number} 期限日数
 */
function getDeadlineDays_() {
  var properties = PropertiesService.getScriptProperties();
  var days = properties.getProperty('DEFAULT_DEADLINE_DAYS');
  return days ? parseInt(days, 10) : 7;
}

/**
 * メール通知用のScript Propertiesを設定する（初回セットアップ用）
 * @param {string} email - 通知先メールアドレス
 */
function setNotificationEmail(email) {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty('NOTIFICATION_EMAIL', email);
  Logger.log('通知メール設定完了: ' + email);
}

// =============================================================================
// データ処理モジュール
// =============================================================================

/**
 * フォーム回答からメタ項目とチェック項目を分離抽出する
 * @param {Object} namedValues - フォームの回答データ
 * @returns {Object} {metaData: Object, checkItems: Object}
 */
function extractFormData(namedValues) {
  // メタ項目集合
  var metaItems = ['設備ID', '点検者', '点検日', '備考', 'タイムスタンプ', 'Timestamp'];
  
  var metaData = {};
  var checkItems = {};
  
  if (!namedValues) {
    Logger.log('警告: namedValuesが空です');
    return {metaData: metaData, checkItems: checkItems};
  }
  
  // 全ての回答項目を分類
  for (var key in namedValues) {
    var value = namedValues[key] ? namedValues[key][0] : '';
    
    if (metaItems.indexOf(key) >= 0) {
      metaData[key] = value;
    } else {
      checkItems[key] = value;
    }
  }
  
  // 必須項目のデフォルト値設定
  metaData['設備ID'] = metaData['設備ID'] || '';
  metaData['点検者'] = metaData['点検者'] || '不明';
  metaData['点検日'] = metaData['点検日'] || formatDate_(new Date());
  metaData['備考'] = metaData['備考'] || '';
  
  return {
    metaData: metaData,
    checkItems: checkItems
  };
}

/**
 * 回答値のNG判定を行う
 * @param {string} answer - 回答値
 * @returns {boolean} true: NG, false: OK
 */
function isNonCompliant(answer) {
  if (!answer || typeof answer !== 'string') {
    return false;
  }
  
  var lowerAnswer = answer.toLowerCase();
  
  // NG判定条件: 「不適合」を含む、または完全一致で「ng」「×」「x」
  return lowerAnswer.indexOf('不適合') >= 0 || 
         lowerAnswer === 'ng' || 
         lowerAnswer === '×' || 
         lowerAnswer === 'x';
}

// =============================================================================
// Issues管理モジュール
// =============================================================================

/**
 * issuesシートの存在確認・新規作成を行う
 * @param {Spreadsheet} spreadsheet - 対象スプレッドシート
 * @returns {Sheet} issuesシート
 */
function createOrGetIssuesSheet(spreadsheet) {
  var sheetName = 'issues';
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    // 新規シート作成
    sheet = spreadsheet.insertSheet(sheetName);
    
    // ヘッダ行を追加（1回のみ）
    var headers = [
      'ID', '状態', '設備ID', '点検日', '点検者',
      '不適合項目', '備考', '起票日', '期限', '完了日', '回答キー'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダ行の書式設定
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');
    
    Logger.log('issuesシートを新規作成しました');
  } else {
    // 既存シートで回答キー列が足りない場合の補修
    if (sheet.getLastColumn() === 10) {
      sheet.getRange(1, 11).setValue('回答キー');
      Logger.log('既存issuesシートに回答キー列を追加しました');
    }
  }
  
  return sheet;
}

/**
 * 既存issuesの重複チェックを行う
 * @param {Sheet} sheet - issuesシート
 * @param {string} uniqueKey - 一意キー
 * @returns {boolean} true: 重複あり, false: 重複なし
 */
function checkDuplicateIssue_(sheet, uniqueKey) {
  if (!uniqueKey) {
    return false;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false; // データ行なし
  }
  
  // 回答キー列での重複チェック（新方式）
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var keyCol = header.indexOf('回答キー') + 1;
  
  if (keyCol > 0) {
    // 回答キー列が存在する場合
    var keys = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues().map(function(r) { return r[0]; });
    return keys.indexOf(uniqueKey) !== -1;
  }
  
  // フォールバック（古いシート用）：従来ロジック
  var rows = sheet.getRange(2, 1, lastRow - 1, Math.min(10, sheet.getLastColumn())).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].length >= 8) {
      var legacyKey = generateUniqueKey_(rows[i][7], rows[i][2]); // 起票日 + 設備ID
      if (legacyKey === uniqueKey) {
        return true;
      }
    }
  }
  
  return false;
}

/**
 * issuesシートに新しい行データを追加する
 * @param {Sheet} sheet - issuesシート
 * @param {Object} formData - フォームデータ
 * @param {Array} nonCompliantItems - 不適合項目の配列
 * @param {string} uniqueKey - 回答キー（重複防止用）
 */
function addIssueRecord(sheet, formData, nonCompliantItems, uniqueKey) {
  var issueDate = formatDate_(new Date());
  var deadlineDays = getDeadlineDays_();
  var deadline = formatDate_(addDays_(new Date(), deadlineDays));
  
  var rowData = [
    'ISS-' + generateUuid_(),                           // ID
    '新規',                                             // 状態
    formData.metaData['設備ID'],                        // 設備ID
    formData.metaData['点検日'],                        // 点検日
    formData.metaData['点検者'],                        // 点検者
    nonCompliantItems.join(' / '),                      // 不適合項目
    formData.metaData['備考'],                          // 備考
    issueDate,                                          // 起票日
    deadline,                                           // 期限
    '',                                                 // 完了日（空白）
    uniqueKey || ''                                     // 回答キー（重複防止）
  ];
  
  // 最下行に追加
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
  
  Logger.log('issues行追加完了: ID=' + rowData[0]);
}

// =============================================================================
// 通知モジュール
// =============================================================================

/**
 * メール通知を送信する（NOTIFICATION_EMAIL設定時のみ）
 * @param {Object} formData - フォームデータ
 * @param {Array} nonCompliantItems - 不適合項目の配列
 * @param {string} sheetUrl - issuesシートのURL
 */
function sendNotification(formData, nonCompliantItems, sheetUrl) {
  var properties = PropertiesService.getScriptProperties();
  var email = properties.getProperty('NOTIFICATION_EMAIL');
  
  if (!email) {
    Logger.log('NOTIFICATION_EMAIL未設定のため、メール送信をスキップします');
    return;
  }
  
  try {
    var subject = '【始業前点検】不適合項目検出 - ' + formData.metaData['設備ID'];
    
    var body = '始業前点検で不適合項目が検出されました。\n\n' +
               '設備ID: ' + formData.metaData['設備ID'] + '\n' +
               '点検者: ' + formData.metaData['点検者'] + '\n' +
               '点検日: ' + formData.metaData['点検日'] + '\n' +
               '不適合項目: ' + nonCompliantItems.join(', ') + '\n\n' +
               '詳細はissuesシートをご確認ください:\n' + sheetUrl + '\n\n' +
               '※このメールは自動送信です。';
    
    MailApp.sendEmail(email, subject, body);
    Logger.log('メール通知送信完了: ' + email);
    
  } catch (error) {
    Logger.log('メール送信エラー: ' + error.message);
  }
}

// =============================================================================
// メイン処理
// =============================================================================

/**
 * フォーム送信イベントのメインハンドラー
 * @param {Event} e - フォーム送信イベント
 */
function onFormSubmit(e) {
  try {
    Logger.log('フォーム送信処理開始');
    
    // 入力検証
    if (!e || !e.namedValues) {
      Logger.log('エラー: イベントデータが不正です');
      return;
    }
    
    // データ抽出
    var formData = extractFormData(e.namedValues);
    
    // 設備ID欠損チェック
    if (!formData.metaData['設備ID']) {
      Logger.log('警告: 設備IDが欠損しているため処理を中止します');
      return;
    }
    
    // 重複チェック用キー生成
    var timestamp = formData.metaData['タイムスタンプ'] || formData.metaData['Timestamp'];
    var uniqueKey = generateUniqueKey_(timestamp, formData.metaData['設備ID']);
    
    if (!uniqueKey) {
      Logger.log('警告: 一意キーが生成できませんでした');
      return;
    }
    
    // スプレッドシート取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var issuesSheet = createOrGetIssuesSheet(spreadsheet);
    
    // 重複チェック
    if (checkDuplicateIssue_(issuesSheet, uniqueKey)) {
      Logger.log('重複キー検出のため処理をスキップします: ' + uniqueKey);
      return;
    }
    
    // NG項目の検出
    var nonCompliantItems = [];
    for (var key in formData.checkItems) {
      if (isNonCompliant(formData.checkItems[key])) {
        nonCompliantItems.push(key);
      }
    }
    
    // NG項目がない場合は正常終了
    if (nonCompliantItems.length === 0) {
      Logger.log('適合: 不適合項目なし');
      return;
    }
    
    // issues行追加
    addIssueRecord(issuesSheet, formData, nonCompliantItems, uniqueKey);
    
    // メール通知送信
    var sheetUrl = spreadsheet.getUrl() + '#gid=' + issuesSheet.getSheetId();
    sendNotification(formData, nonCompliantItems, sheetUrl);
    
    Logger.log('フォーム送信処理完了: NG項目数=' + nonCompliantItems.length);
    
  } catch (error) {
    Logger.log('予期せぬエラー: ' + error.message);
  }
}

// =============================================================================
// 運用支援機能
// =============================================================================

/**
 * 状態が「完了」に変更されたときに完了日を自動記録する
 * @param {Event} e - セル編集イベント
 */
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'issues') {
      return;
    }
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colStatus = headers.indexOf('状態') + 1;
    var colDone = headers.indexOf('完了日') + 1;
    
    // 状態列の編集でない場合、またはヘッダ行の場合はスキップ
    if (e.range.getColumn() !== colStatus || e.range.getRow() === 1) {
      return;
    }
    
    var value = (e.value || '').toString().trim();
    if (value === '完了' && colDone > 0) {
      // 完了日を当日で自動設定
      var today = formatDate_(new Date());
      sheet.getRange(e.range.getRow(), colDone).setValue(today);
      Logger.log('完了日自動記録: ' + today + ' (行' + e.range.getRow() + ')');
    }
  } catch (error) {
    Logger.log('onEdit エラー: ' + error.message);
  }
}