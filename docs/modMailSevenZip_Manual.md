# Outlook Mail Saver with 7-Zip Auto-Extraction (modMailSevenZip)

これは、OutlookのVBAマクロです。Outlookで選択したメールを「ダウンロード」フォルダに保存し、添付ファイルを自動解凍します。また、添付ファイルが暗号化ZIPまたは7z形式だった場合に、事前定義したパスワードリスト（日付変数対応）を用いて自動解凍を試みます。

## 特徴

* **メール保存**: 選択したメールの添付ファイルと本文（UTF-8テキスト）を `Downloads\yymmdd_hhmmss_件名` フォルダに一括保存します。
* **自動解凍**: 添付ファイルが `.zip` または `.7z` の場合、インストール済みの **7-Zip** を使用して解凍を試みます。
* **分割Zip自動結合**: 複数メールを選択した状態で実行すると、複数通に分かれて送られてきた分割アーカイブ（`.001`, `.002`, ...）を自動的に1つのフォルダに集約・結合して解凍します。
* **パスワード総当たり**: 設定ファイルに基づき、日付（`{mm}{dd}`など）を含むパスワード候補を自動生成して解凍をテストします（行頭の `//` はコメント、`#` 始まりのパスワードや `"` 囲みにも対応）。
* **共通ログ機能**: 共通モジュール `modLogger` を使用し、実行ログをローカルファイルに記録します。
* **手動フォールバック**: 自動解凍に失敗した場合、インプットボックスを表示して手動でパスワード入力を求めます（無限リトライ可）。
* **安全性**: 保存先フォルダが既に存在する場合は、上書きを行わずにそのフォルダを開きます（重複保存防止）。なお、フォルダ内の添付ファイル保存や解凍処理については、同名ファイルがある場合に連番を付与して上書きを回避します。

## 必要要件

* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* [7-Zip](https://7-zip.opensource.jp/)
    * デフォルトインストールパス: `C:\Program Files\7-Zip\7z.exe`
    * ※異なる場所にインストールしている場合は `config.ini` で指定可能です。
* **共通モジュール**:
    * `src/modConfig.bas` (共通設定モジュール)
    * `src/modLogger.bas` (共通ログモジュール)

## インストール

1. このリポジトリのファイルをダウンロードします。
2. Outlookを起動し、`Alt + F11` でVBAエディタを開きます。
3. **モジュールのインポート**:
    * `File` > `Import File` から以下の**3つのファイル**をインポートします。
        1. `src/modConfig.bas` (必須：設定読み込み用)
        2. `src/modLogger.bas` (必須：ログ出力用)
        3. `src/modMailSevenZip.bas` (メイン機能モジュール)


4. **ThisOutlookSessionの設定**:
    * VBAエディタ左側の `Project1` > `Microsoft Outlook Objects` > `ThisOutlookSession` をダブルクリックします。
    * 同梱の `src/ThisOutlookSession.cls` の中身をコピーし、貼り付けます。
    * または、以下のプロシージャを追加してください。
       ```vb
       ' 選択したメールの添付ファイルを保存・解凍するマクロ
       Public Sub メールをダウンロードフォルダに保存()
           ' modMailSevenZip内でRunID生成・ログ出力まで完結するため、直接呼び出します
           modMailSevenZip.SaveAndExtractAttachments
       End Sub
       ```

## 設定（パスワードリスト）

解凍時に試行するパスワードリストを設定します。
セキュリティ向上のため、**PowerShellスクリプトによる暗号化保管（推奨）** と **平文ファイル（下位互換）** の2つの方式に対応しています。

### 方式A: PowerShell による暗号化保管（強く推奨）

Windows DPAPI (Data Protection API) を利用し、ログインユーザーのアカウントに紐づけてパスワードを暗号化保存します。
ファイル（`SevenZipPasswords.enc`）を他人が取得しても復号できないため、平文保存のリスクを完全に解消できます。管理者権限のない標準ユーザーでも実行可能です。

#### 1. パスワードの新規登録
PowerShell（Windows TerminalまたはPowerShell 5.1/7）で以下を実行します。

```powershell
# 1件または複数件のパスワードを暗号化して保存
.\scripts\Protect-SevenZipPassword.ps1 -Password "MyPassword123", "Secret#{yyyy}!"
```

* 既存の暗号化ファイルに追記したい場合は `-Append` スイッチを付与します：
  ```powershell
  .\scripts\Protect-SevenZipPassword.ps1 -Password "AdditionalPass" -Append
  ```

#### 2. 既存の平文ファイルからの移行
すでに `SevenZipPasswords.txt` をお持ちの場合は、一括で暗号化ファイルに変換できます：

```powershell
# 平文ファイルから一括インポートして暗号化
.\scripts\Protect-SevenZipPassword.ps1 -ImportFromTxt

# （移行後に平文の元ファイルを自動削除したい場合）
.\scripts\Protect-SevenZipPassword.ps1 -ImportFromTxt -RemoveSource
```

#### 3. 登録パスワードの確認
登録されたパスワード一覧を確認したい場合は、`Unprotect-SevenZipPassword.ps1` を実行します：

```powershell
# マスク表示（安全確認用）
.\scripts\Unprotect-SevenZipPassword.ps1

# 平文表示（内容確認用）
.\scripts\Unprotect-SevenZipPassword.ps1 -ShowPlain
```

---

### 方式B: 平文テキストファイル（下位互換用）

スクリプトによる暗号化を行わず、従来のテキストファイル（`SevenZipPasswords.txt`）をそのまま使用することも可能です。
（※暗号化ファイル `SevenZipPasswords.enc` が存在しない場合に自動的に読み込まれます）

1. **ファイルの準備**:
    * `configs/SevenZipPasswords.sample.txt` をコピーして、名前を `SevenZipPasswords.txt` に変更します。
2. **ファイルの配置**:
    * エクスプローラーで `%APPDATA%\OutlookVBA\` を開きます。
    * 用意した `SevenZipPasswords.txt` を配置します。
3. **リストの編集**:
    * テキストエディタで開き、よく使うパスワードパターンを記述します（文字コードは **UTF-8**）。

---

### テンプレート変数（両方式共通）

パスワード内に以下のプレースホルダーを含めることで、**メールの受信日** に応じたパスワードを自動生成できます：

* `{yyyy}`: 西暦4桁 (例: 2026)
* `{yy}`: 西暦下2桁 (例: 26)
* `{mm}`: 月2桁 (例: 09)
* `{dd}`: 日2桁 (例: 10)

```text
// 記述例
mytag{mm}{dd}
#pass_{yyyy}
fixed_password
" pass with space "
```
* **記号対応**: 行頭が `#` のパスワードや、前後に空白がある場合は `"` で囲んで記述可能です。コメント行は `//` で始めます。
## 設定（7-Zipパスなど）

7-Zipのインストール先が標準と異なる場合などは、共通設定ファイル `config.ini` を編集します。

1. `%APPDATA%\OutlookVBA\config.ini` を開きます（存在しない場合は `configs/config.sample.ini` をコピーして配置）。
2. `[General]` セクションを編集します。
    ```ini
    [General]
    # 7-Zipの実行ファイルパス
    SevenZipPath=D:\Tools\7-Zip\7z.exe
    ```

## 使い方

1. Outlookで保存したいメールを選択します（複数選択時は最初の1件が対象）。
2. マクロ `メールをダウンロードフォルダに保存` を実行します。
    * リボンやクイックアクセスツールバーにボタンを登録しておくと便利です。


3. 処理が完了すると、保存先のフォルダがエクスプローラで自動的に開きます。
    * 暗号化ZIPが含まれていた場合、解凍済みのフォルダも作成されています。
    * パスワードがリストにない場合、入力ダイアログが表示されます。

## ログ

本ツールは `modLogger` を使用して動作ログを記録します。
エラー発生時の調査や、どのパスワードで解凍に成功したかを確認する場合に参照してください。

* **ログ保存場所**: `%APPDATA%\OutlookVBA\logs\`
* **ファイル名**: `yyyy-mm-dd.log` (日付ごとのローテーション)
* **文字コード**: UTF-8
* **ログ出力例:**
    ```text
    2026/02/16 12:00:00.123 [260216-120000-SAVE] [MailSevenZip] === START メール保存・解凍処理 ===
    2026/02/16 12:00:00.250 [260216-120000-SAVE] [MailSevenZip] 対象メール：Subject="請求書送付" / Received=2026/02/16 10:00:00
    ...
    2026/02/16 12:00:05.500 [260216-120000-SAVE] [MailSevenZip] テスト成功："pass_0216" → 抽出へ
    ```

## ライセンス

MIT License