# Outlook Auto Flag (メール自動フラグ付与ツール)

これは、OutlookのVBAマクロです。受信したメール（および起動時の未読メール）を解析し、設定ファイルで定義した条件（送信者、本文キーワード、件名除外）に合致する場合に、自動的に「要確認」フラグを設定します。

## 特徴

* **自動フラグ付与**: 条件に合致したメールに「要確認」フラグ (`olFlagMarked`) を即座に設定します。
* **高度な条件判定**:
* **宛先チェック**: To または CC に自分の名前やアドレスが含まれているか（重要メールの判定）。
* **正規表現マッチ**: 本文に含まれるキーワードを正規表現（例: `緊急|重要|回答期限`）で柔軟に指定可能。
* **件名除外**: 「自動通知」「日報」など、フラグ不要なメールを件名キーワードで除外。


* **外部設定ファイル**: 設定値は `%APPDATA%` 配下の統合設定ファイル (`config.ini`) で管理します。VBAコードを編集せずに条件を変更できます。
* **起動時チェック**: Outlook起動時に、受信トレイにある「未読メール」を遡ってチェックします。
* **統一ログ管理**: 共通ログモジュール (`modLogger`) を使用し、詳細な実行履歴を記録します。
* **安全性**: 会議出席依頼やタスク依頼など、メール以外のアイテムは処理対象外としてスキップします。

## 必要要件

* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* **必須モジュール**: `modLogger.bas` (本リポジトリに同梱されている共通ログモジュール)

## インストール

1. このリポジトリのファイルをダウンロードします。
2. Outlookを起動し、`Alt + F11` でVBAエディタを開きます。
3. **モジュールのインポート**:
    * `File` > `Import File` から **`src/modAutoFlag.bas`** をインポートします。
    * 続けて、**`src/modLogger.bas`** もインポートします（※本ツールは `modLogger` に依存しているため、必ず両方インポートしてください）。


4. **ThisOutlookSessionの設定**:
    * VBAエディタ左側の `Project1` > `Microsoft Outlook Objects` > `ThisOutlookSession` をダブルクリックします。
    * 以下のコードを貼り付けます（`src/ThisOutlookSession.cls` を参考にしてください）。


        ```vb
        Private Sub Application_Startup()
            ' 起動時に未読メールをチェック
            modAutoFlag.ProcessStartupUnread
        End Sub

        Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
            ' 新着メール受信時にチェック
            modAutoFlag.ProcessNewMail EntryIDCollection
        End Sub

        ```



## 設定（config.ini）

本ツールは、共通設定ファイル `%APPDATA%\OutlookVBA\config.ini` の **`[AutoFlag]`** セクションを使用します。

1. エクスプローラーで `%APPDATA%\OutlookVBA\` を開きます。
2. `config.ini` をテキストエディタ（メモ帳など）で開きます。
    * ※ファイルがない場合は、`configs/config.sample.ini` を参考に作成してください。


3. **`[AutoFlag]`** セクションを編集します。
    * **重要**: 保存時の文字コードは必ず **UTF-8 (BOMなし推奨)** にしてください。Shift-JISでは文字化けします。


        ```ini
        [AutoFlag]
        # 行頭に「#」をつけるとコメントになります。

        # [MyAddress]
        # 自分のメールアドレス、またはOutlookでの表示名（例: 山田 太郎）
        # ToまたはCCにこの文字列が含まれるメールが処理対象になります。
        MyAddress=taro.yamada@example.com

        # [Pattern]
        # 本文に含まれるとフラグを立てるキーワード（正規表現）
        # 複数のキーワードはいずれか(|)で区切ります。
        Pattern=重要|緊急|要回答|期限

        # [ExcludeSubjects]
        # 処理から除外したい件名のキーワード（カンマ区切り）
        # 自動通知メールなどを除外するのに便利です。
        ExcludeSubjects=さんがメッセージを送信しました,日報,自動通知

        ```



## 使い方

1. 設定ファイルの配置とVBAの導入が完了したら、Outlookを再起動します。
2. **自動実行**:
    * 以降、メールを受信するたびに自動的に判定が行われます。
    * Outlook起動時にも、未読メールに対して判定が行われます。


3. 条件に一致したメールには、自動的に赤い「フラグ」が付きます。

## ログ

実行時の動作ログは、以下の2箇所に出力されます。動作確認やトラブルシューティングにご利用ください。

1. **イミディエイト ウィンドウ**: VBAエディタ内で `Ctrl + G` を押すと確認できます。
2. **ログファイル**: `%APPDATA%\OutlookVBA\logs\` フォルダに日付ごとのログファイルが生成されます。

    **ログ出力例:**

    ```text
    2026/02/15 09:00:00.123 [260215-090000-BOOT] [AutoFlag] === START メール自動フラグ処理 (起動時未読チェック) ===
    2026/02/15 09:00:00.150 [260215-090000-BOOT] [AutoFlag] 設定準備完了: Pattern=重要|緊急, Excludes=3件
    2026/02/15 09:00:01.005 [260215-090000-BOOT] [AutoFlag] >>> フラグ設定: プロジェクト進捗について
    2026/02/15 09:00:01.020 [260215-090000-BOOT] [AutoFlag] === END メール自動フラグ処理 (起動時未読チェック) ===

    ```

## ライセンス

MIT License