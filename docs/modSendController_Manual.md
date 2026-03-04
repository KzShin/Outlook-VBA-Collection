# Outlook Mail Send Controller (modSendController)

これは、OutlookのVBAマクロです。メール送信ボタンが押されたタイミングで割り込み処理を行い、誤送信防止（添付忘れ、セキュリティチェック）および送信日時の自動制御（時間外・休日送信の予約）を行います。

**本モジュールは、共通ログモジュール `modLogger` に依存します。**

## 特徴

送信フローは以下の3ステップで構成されています。

1. **添付忘れ防止 (Step 1)**:
    * 本文に「添付」というキーワードが含まれているにもかかわらず、ファイルが添付されていない場合に警告ダイアログを表示します。


2. **Zip/7z 暗号化チェック (Step 2)**:
    * 添付ファイルに `.zip` または `.7z` が含まれる場合、外部ツールの **7-Zip** を使用して暗号化（パスワード保護）の有無を判定します。
    * パスワードがかかっていない圧縮ファイルが見つかった場合、警告を表示して送信を中止できます。
    * 同名の添付ファイルが複数あっても、連番処理により正確に判定します。


3. **送信時刻・休日制御 (Step 3)**:
    * **夜間・早朝**: 設定された業務時間外（デフォルトは 18:00～翌07:59）の送信操作に対し、翌営業日の業務開始時刻（デフォルトは 08:00）への予約送信を提案します。
    * **休日・祝日**: 土日および設定ファイルで定義した祝日の場合、翌営業日の業務開始時刻への予約送信を提案します。
    * **予約機能**: ダイアログで「Yes」を選択すると、Outlookの「配信タイミング」機能を自動設定して送信トレイに待機させます。

## 必要要件

* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* [7-Zip](https://7-zip.opensource.jp/)
* 暗号化チェック機能を使用するために必須です。
* デフォルトパス: `C:\Program Files\7-Zip\7z.exe` (設定ファイルで変更可能)


* **共通モジュール**: `modLogger.bas` (同梱のログ出力用モジュール)

## インストール

1. このリポジトリのファイルをダウンロードします。
2. Outlookを起動し、`Alt + F11` でVBAエディタを開きます。
3. **モジュールのインポート**:
    * `File` > `Import File` から以下の**2つのファイル**をインポートします。
        1. `src/modLogger.bas` (必須：ログ出力用)
        2. `src/modSendController.bas` (本体)


4. **ThisOutlookSessionの設定**:
    * VBAエディタ左側の `Project1` > `Microsoft Outlook Objects` > `ThisOutlookSession` をダブルクリックします。
    * 以下のコードを記述（または同梱の `src/ThisOutlookSession.cls` からコピー）します。
       ```vb
       Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
           ' コントローラーへ処理を委譲
           modSendController.Execute Item, Cancel
       End Sub    
       ```

## 設定 (Configuration)

本ツールは、共通設定ファイル `%APPDATA%\OutlookVBA\config.ini` の **`[SendController]`** および **`[General]`** セクションを使用します。

1. **設定ファイルの準備**:
    * エクスプローラーで `%APPDATA%\OutlookVBA\` を開きます。
    * `config.ini` をテキストエディタで開きます（ファイルがない場合は `configs/config.sample.ini` をコピーして作成）。


2. **設定の編集**:
    * 以下のセクションを確認・編集します。
    * **重要**: 保存時の文字コードは必ず **UTF-8 (BOMなし推奨)** にしてください。
      ```ini
      [General]
      # [共通] 7-Zipの実行ファイルパス
      # 7-Zipがデフォルト以外の場所にインストールされている場合のみ変更してください。
      SevenZipPath=C:\Program Files\7-Zip\7z.exe
      
      [SendController]
      # [誤送信防止] 送信保留する祝日リスト (MM-DD形式, カンマ区切り)
      # 土日に加えて、ここで指定した日付も「休日」とみなし、翌営業日送信を提案します。
      # 例: 年末年始休暇など
      HolidayList=12-29,12-30,12-31,01-01,01-02,01-03
      
      # 業務開始時間（兼 翌営業日の予約送信時刻）
      WorkStartTime=08:00
      
      # 業務終了時間（この時間以降の送信は翌営業日扱いで提案）
      WorkEndTime=18:00    
      ```

## 使い方

1. Outlookでメールを作成し、「送信」ボタンを押します。
2. マクロが自動的に起動し、以下のチェックを行います。
    * **警告**: 問題（添付忘れ、パスワードなしZip）がある場合、警告ダイアログが出ます。「いいえ」を選ぶと送信画面に戻ります。
    * **確認**: 営業時間外の場合、予約送信するかどうかの確認ダイアログが出ます。「Yes」で予約、「No」で即時送信、「Cancel」で送信中止となります。


3. 問題がなければ、そのまま送信（または送信トレイへ予約保存）されます。

## ログ (Logging)

本ツールは `modLogger` を介してログを出力します。

* **イミディエイトウィンドウ**: VBAエディタの `Ctrl + G` でリアルタイムに確認できます。
* **ログファイル**: 以下のフォルダに日次で保存されます。
* 保存先: `%APPDATA%\OutlookVBA\logs\yyyy-mm-dd.log`
* 実行ID (`RunId`) により、一連の処理フローを追跡可能です。

    **ログ出力例:**
    ```text
    2026/02/16 20:00:00.123 [260216-200000-SEND] [SendController] === START SendController ===
    2026/02/16 20:00:00.125 [260216-200000-SEND] [SendController] Subject=テストメール / Attachments=1
    2026/02/16 20:00:00.130 [260216-200000-SEND] [SendController] Step1: Skipped.
    2026/02/16 20:00:00.250 [260216-200000-SEND] [SendController] Step2 Target: ...\chk_200000_1_data.zip
    2026/02/16 20:00:00.500 [260216-200000-SEND] [SendController] Step3 Status: Night/Early. Candidate=2026/02/17 08:00:00
    2026/02/16 20:00:05.000 [260216-200000-SEND] [SendController] Step3 User Selection: Yes
    2026/02/16 20:00:05.010 [260216-200000-SEND] [SendController] === END SendController / Cancel=False ===    
    ```
## ライセンス
MIT License
