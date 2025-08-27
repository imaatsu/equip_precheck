# 設計書

## アーキテクチャ概要
Google Apps Scriptによるイベントドリブン設計
- フォーム送信 → onFormSubmit(e)自動実行 → データ検証・処理 → issuesシート更新 → メール通知

## モジュール設計

### 1. メイン処理モジュール
**onFormSubmit(e)**
- フォーム送信イベントのメインハンドラー
- 入力値検証、NG判定、issues起票、メール通知を統括

### 2. データ処理モジュール
**extractFormData(namedValues)**
- namedValuesからメタ項目とチェック項目を分離抽出
- メタ項目集合: ['設備ID','点検者','点検日','備考','タイムスタンプ','Timestamp']

**isNonCompliant(answer)**
- 回答値のNG判定ロジック
- 判定条件: 小文字変換後「不適合」を含む、または完全一致で「ng」「×」「x」

### 3. Issues管理モジュール
**createOrGetIssuesSheet(spreadsheet)**
- issuesシートの存在確認・新規作成
- ヘッダ行の1回限定追加

**addIssueRecord(sheet, formData, nonCompliantItems)**
- issues行データの生成・追加
- UUID生成、期限計算、データ整形

### 4. ユーティリティモジュール
**addDays_(date, days)**
- 日付加算処理（期限計算用）

**formatDate_(date)**
- yyyy-MM-dd形式の日付文字列変換

**generateUuid_()**
- 8桁UUID断片生成（ID用）

**generateUniqueKey_(timestamp, equipmentId)**
- 重複防止用一意キー生成

**checkDuplicateIssue_(sheet, uniqueKey)**
- 既存issuesの重複チェック

**setNotificationEmail(email)**
- Script Properties設定ユーティリティ

**getDeadlineDays_()**
- DEFAULT_DEADLINE_DAYS取得（未設定時7）

### 5. 通知モジュール
**sendNotification(formData, nonCompliantItems, sheetUrl)**
- メール通知送信処理
- NOTIFICATION_EMAIL設定時のみ実行
- issuesシートURLを本文に含める

## データフロー
1. フォーム送信 → e.namedValues取得
2. メタ項目/チェック項目分離 → extractFormData()
3. 設備ID欠損チェック（欠損時は警告ログで終了）
4. 重複キー生成・チェック → generateUniqueKey_() + checkDuplicateIssue_()
5. 各チェック項目のNG判定 → isNonCompliant()
6. NG項目がある場合（かつ非重複）:
   - issuesシート準備 → createOrGetIssuesSheet()
   - issues行追加 → addIssueRecord()
   - メール通知 → sendNotification()
7. 全て適合または重複の場合: 処理終了

## エラーハンドリング戦略
- **入力検証**: e/e.namedValues存在チェック
- **Null安全**: メタ項目取得時のデフォルト値設定
- **例外捕捉**: メール送信失敗時のログ出力
- **冪等性**: issuesシートヘッダの重複追加防止
- **重複防止**: 一意キーで既存issuesの重複チェック
- **ログセキュリティ**: 個人情報（備考等）はログ出力しない

## パフォーマンス考慮事項
- スプレッドシート操作の最小化（バッチ処理）
- Rangeに配列で一括書き込み（setValues使用）
- 文字列操作の効率化（toLowerCase()の1回実行）
- 不要なAPI呼び出し回避（条件分岐による最適化）

## セキュリティ設計
- Script Properties使用による設定情報の保護
- 入力値のサニタイゼーション（メール内容）
- 権限最小化（必要最低限のスコープ）

## 拡張性考慮
- チェック項目の動的検出（メタ項目集合による判定）
- 設問追加時のコード変更不要設計
- 設定パラメータの外部化（Script Properties）

## 実装上の注意点
- タイムゾーン固定（Asia/Tokyo）
- 日本語コメント必須
- 1ファイル完結
- 省略記法禁止
- JSDoc風説明付与
- インストール型トリガー設定が必要

## テスト観点
- 正常系: 適合/不適合の各パターン
- 異常系: 不正な入力値、シート操作エラー
- 境界値: メタ項目の欠損、空文字列
- 設定系: メール通知ON/OFF
- 重複系: 回答編集後の再発火で重複起票しない
- 欠損系: 設備ID欠損時は起票せず警告のみ