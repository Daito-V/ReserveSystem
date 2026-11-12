# 会議室予約システム（GAS）

要件定義書 v1.0 の今回対象（F-1〜F-5）を、Google Apps Script と Google スプレッドシートで実装したプロジェクトです。

## ローカルで編集・GitHub管理

1. Node.js を用意し、`npm install -g @google/clasp` を実行します。
2. `clasp login` でGoogleアカウントにログインします。
3. Apps Script の新規プロジェクトを作成し、プロジェクトのスクリプトIDを取得します。
4. `clasp clone <スクリプトID>` でこの作業ディレクトリに取得するか、`.clasp.json` を `{ "scriptId": "取得したID", "rootDir": "." }` として作成します。
5. `clasp push` でソースをApps Scriptへ反映します。
6. 対象のスプレッドシートに紐づくスクリプトとして使う場合、スプレッドシートからApps Scriptを作成してスクリプトIDを使ってください。

`.clasp.json` はGoogleプロジェクト固有なのでGitHubには登録しません。GitHubには本ファイル群と `appsscript.json` を保存します。

## 初期設定と起動

1. Apps Script エディタで `setup` を一度実行し、権限を承認します。`rooms`、`users`、`reservations`、`cancellations` シートと列見出しを作り、会議室5室を登録します。
2. `users` シートに社員ID、氏名、`ACTIVE` / `LEAVE` / `RETIRED`、`MEMBER` / `ADMIN` を登録します。利用者IDは画面上で選択します。認証はありません。
3. `clasp push` で最新ファイルを反映し、Apps Scriptの「デプロイ」からウェブアプリとしてデプロイします。

## 実装範囲

- 日付別予約一覧、会議室別・開始時刻順の表示、キャンセル済み予約の取り消し線表示
- 新規予約、同一会議室の時間重複防止、社員ごとの1日3件制限
- 本人キャンセル（開始24時間前まで）と総務部管理者の代理キャンセル
- キャンセル実行者・日時・本人/代理区分・理由の記録と詳細表示
- 登録処理とキャンセル処理を `LockService` で直列化

仕様書の今回対象外に従い、予約変更、入退室記録、自動キャンセル、利用率レポート、ログイン認証、通知は実装していません。Vercelは使用せず、GASのウェブアプリとして実行します。
