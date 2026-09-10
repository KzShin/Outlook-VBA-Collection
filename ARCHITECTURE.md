# Outlook VBA Collection - Architecture

本書は、`Outlook-VBA-Collection` リポジトリの全体的なアーキテクチャ、コンポーネント間の関係、およびデータフローについて説明します。

## 1. システム概要

本システムは、Microsoft Outlook 上で動作する VBA (Visual Basic for Applications) のスクリプト群です。
Outlook の標準機能を拡張し、誤送信の防止、添付ファイルの処理自動化、定型業務の効率化を実現します。
各機能は独立したモジュールとして実装されていますが、共通の「設定管理」と「ログ管理」基盤に依存することで、保守性と拡張性を高めています。

## 2. ディレクトリ構造と主要コンポーネント

リポジトリは主に以下のディレクトリで構成されています。

```text
Outlook-VBA-Collection/
├── src/        : VBAソースコード (.bas, .cls, .frm)
├── scripts/    : PowerShellユーティリティスクリプト (.ps1)
├── configs/    : 設定ファイルのサンプル (.ini, .txt)
└── docs/       : ドキュメント群
```

### 2.1 エントリポイント
* **`ThisOutlookSession.cls`**
  Outlook のアプリケーションイベント（`Application_ItemSend`, `Application_NewMailEx` など）を捕捉するエントリポイントです。
  ユーザーの操作（送信ボタン押下など）やイベント発生時に、対応するビジネスロジックモジュール（コントローラ）を呼び出します。

### 2.2 共通基盤モジュール (Infrastructure)
すべてのビジネスロジックモジュールから利用される共通機能です。

* **`modConfig` (設定管理)**
  `%APPDATA%\OutlookVBA\config.ini` から設定値を読み込み、メモリ上にキャッシュします。
  各機能の有効/無効、保存先パスなどのパラメータを一元管理します。
* **`modLogger` (ログ管理)**
  システムの動作ログを記録します。
  出力先は標準で `%APPDATA%\OutlookVBA\logs\` となり、エラーの追跡や監査に利用されます。

### 2.3 ビジネスロジックモジュール (Business Logic)
各業務要件を実現する個別の機能モジュールです。

* **送信制御系**
  * **`modSendController` (誤送信防止コントローラ)**
    `ThisOutlookSession` の `ItemSend` イベントから呼び出され、メール送信前のチェック（添付忘れ、件名なし、Zip暗号化チェック）や、時間外・休日の自動送信遅延（翌営業日配送）のロジックを担います。
* **受信・整理系**
  * **`modAutoFlag` (メール自動フラグ)**
    新着メールを受信した際に起動し、件名や本文を正規表現で解析して自動的にフラグ（重要、緊急など）を付与します。
  * **`modMailOpen` (メール開封制御)**
    意図しないダブルクリック等によるメールの別ウィンドウ表示を防ぎ、特定の操作時のみ開封を許可する制御を行います。
* **添付ファイル操作系**
  * **`modMailSevenZip` (添付ファイル保存＆解凍)**
    選択したメールの添付ファイルを抽出し、7-Zip を用いて自動解凍・保存を行います。パスワード付きZipの場合は外部リストからパスワードを試行します。
* **作成・出力系**
  * **`modTemplateMail` (テンプレートメール作成)**
    定型文テンプレートを読み込み、日付や宛名などの変数を置換して新規メールを作成します（`frmSelectTemplate` と連携）。
  * **`modForwardDraft` (転送下書き作成)**
    受信メールを指定の宛先・定型文を付与した状態のテキストメール下書きとして生成します（`frmSelectDest` と連携）。
  * **`modMailToPDF` (PDF出力)**
    選択されたメールをPDFプリンタドライバ経由でPDFファイルとして出力・保存します。
  * **`modLaunchOWA` (OWA起動)**
    Outlook on the Web (OWA) をブラウザで起動します。

### 2.4 ユーザーインターフェース (UI Forms)
ユーザーに入力を求めるためのフォームです。

* **`frmSelectDest.frm / .frx`**: 送信先や転送先を選択するダイアログ。
* **`frmSelectTemplate.frm / .frx`**: 利用するメールテンプレートを選択するダイアログ。

## 3. 外部依存と連携

本システムは、VBA内部に留まらず、いくつかの外部リソースと連携します。

1. **ファイルシステム (設定・ログ・パスワード)**
   * `%APPDATA%\OutlookVBA\` を基準ディレクトリとして、設定ファイル (`config.ini`)、暗号化パスワード (`SevenZipPasswords.enc`)、ログファイルへのアクセスを行います。
2. **外部アプリケーション (7-Zip)**
   * `modSendController` や `modMailSevenZip` において、Zipファイルの解析や解凍のために、外部プロセスとして `7z.exe` (または `7za.exe`) を呼び出します。
3. **PowerShell スクリプト (`scripts/`)**
   * DPAPIを用いたパスワードの暗号化 (`Protect-SevenZipPassword.ps1`) や、ソースコードの文字コード変換 (`Convert-Encoding.ps1`) など、VBA単体では実装が困難または非効率な処理をPowerShellで補完しています。

## 4. データフローの例

**例: メールの送信時 (modSendController)**
1. ユーザーがメールの「送信」ボタンを押下。
2. `ThisOutlookSession.Application_ItemSend` イベントが発火。
3. `modConfig` から設定（チェック機能の有効/無効、時間外設定など）を取得。
4. `modSendController.Execute` が呼び出される。
5. （添付ファイルがある場合）`modSendController` が 7-Zip を呼び出し、暗号化Zipか判定。
6. 問題が見つかった場合、送信をキャンセルしユーザーに警告ダイアログを表示。
7. 処理の全過程は `modLogger` によりログファイルに記録。

## 5. 設計思想 (Design Principles)

* **疎結合**: 各ビジネスロジックは独立しており、不要なモジュールはインポートせずに除外することが可能です（`modConfig` と `modLogger` のみ必須）。
* **設定の外部化**: 定数や閾値（業務時間、キーワード、パス等）はソースコード内にハードコーディングせず、`config.ini` に外部化しています。
* **セキュアなパスワード管理**: パスワードリストは平文ではなく、Windows標準の暗号化基盤(DPAPI)を利用したファイル (`.enc`) に保存することを推奨する設計になっています。
