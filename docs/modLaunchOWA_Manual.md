# Outlook OWA Launcher (modLaunchOWA)

これは、OutlookのVBAマクロです。現在使用しているデスクトップ版Outlookから、**Web版 Outlook (Outlook on the Web / OWA)** を素早く既定のブラウザで開きます。

検索機能の補完や、Web版でしか利用できない機能（特定の会議室予約オプションやアドインなど）にアクセスする際に便利です。

## 特徴

* **ワンクリック起動**: リボンやクイックアクセスツールバーに登録することで、即座にWebメールへアクセスできます。
* **URLカスタマイズ**: 法人向け (Microsoft 365) と 個人向け (Outlook.com) の両方に対応しており、設定ファイルでURLを変更可能です。
* **共通ログ機能**: 共通モジュール `modLogger` を使用し、実行履歴を記録します。

## 必要要件

* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* **共通モジュール**: `modLogger.bas` (本リポジトリに含まれる共通ログモジュール)

## インストール

1. このリポジトリのファイルをダウンロードします。
2. Outlookを起動し、`Alt + F11` でVBAエディタを開きます。
3. **モジュールのインポート**:
    * `File` > `Import File` から以下の2つのファイルをインポートします。
        1. `src/modLogger.bas` (共通ログモジュール)
        2. `src/modLaunchOWA.bas` (メイン機能モジュール)
4. **ThisOutlookSessionの設定**:
    * VBAエディタ左側の `Project1` > `Microsoft Outlook Objects` > `ThisOutlookSession` をダブルクリックします。
    * 同梱の `src/ThisOutlookSession.cls` の中身をコピーし、貼り付けます。
    * または、以下のプロシージャを追加してください。


        ```vb
        ' WEB版Outlook (OWA) を開くマクロ
        Public Sub WEB版Outlook起動()
            ' modLaunchOWA内でRunID生成・ログ出力まで完結しているため、直接呼び出します
            modLaunchOWA.LaunchOWA
        End Sub

        ```



## 設定 (Configuration)

本ツールは、共通設定ファイル `%APPDATA%\OutlookVBA\config.ini` の **`[LaunchOWA]`** セクションを使用します。

1. **設定ファイルの準備**:
    * エクスプローラーで `%APPDATA%\OutlookVBA\` を開きます。
    * `config.ini` をテキストエディタで開きます（ファイルがない場合は `configs/config.sample.ini` をコピーして作成）。


2. **設定の編集**:
    * 使用しているアカウントの種類に合わせて `BaseUrl` を設定します。
    * **重要**: 保存時の文字コードは必ず **UTF-8 (BOMなし推奨)** にしてください。


        ```ini
        [LaunchOWA]
        # Web版OutlookのURL

        # パターンA: 法人用 (Microsoft 365 / Office 365)
        BaseUrl=https://outlook.office.com/mail/

        # パターンB: 個人用 (Outlook.com / Hotmail)
        # BaseUrl=https://outlook.live.com/mail/

        ```



## 使い方

このマクロは、ボタンに割り当てて使用することを推奨します。

1. **ボタンの登録**:
    * Outlookのウインドウ上部で右クリックし、「リボンのユーザー設定」または「クイックアクセスツールバーのユーザー設定」を開きます。
    * 「コマンドの選択」で「マクロ」を選びます。
    * `Project1.ThisOutlookSession.WEB版Outlook起動` を選択し、右側のリストに追加します。
    * 必要に応じて「名前の変更」でアイコンや表示名を変更します（地球儀のアイコンなどがおすすめです）。


2. **実行**:
    * 追加したボタンをクリックすると、既定のブラウザで Web版Outlook が開きます。
    * ※初回アクセス時はログインを求められる場合があります。



## ログ

本ツールは `modLogger` を使用して動作ログを記録します。

* **ログ保存場所**: `%APPDATA%\OutlookVBA\logs\`
* **ファイル名**: `yyyy-mm-dd.log` (日付ごとのローテーション)
* **ログ出力例:**
    ```text
    2026/02/16 09:30:00.123 [260216-093000] [LaunchOWA] === START OWA Launch ===
    2026/02/16 09:30:00.130 [260216-093000] [LaunchOWA] Loaded URL from config: https://outlook.office.com/mail/
    2026/02/16 09:30:00.135 [260216-093000] [LaunchOWA] Opening URL...
    2026/02/16 09:30:00.140 [260216-093000] [LaunchOWA] === END OWA Launch ===

    ```



## ライセンス

MIT License