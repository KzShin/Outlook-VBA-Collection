# Outlook VBA Collection
業務効率化のための Microsoft Outlook 用 VBA スクリプト集です。
誤送信防止、添付ファイルの自動保存・解凍、メール整理の自動化などの機能を提供します。
すべてのモジュールは共通のログ管理モジュール (`modLogger`) を使用して、堅牢かつ追跡可能な動作を実現しています。
## 収録ツール一覧
| モジュール名 | 機能概要 |
| --- | --- |
| **modSendController** | **誤送信防止コントローラ**。<br>添付忘れチェック、Zip暗号化有無の確認、時間外・休日送信の自動予約（翌営業日配信）を行います。 |
| **modMailSevenZip** | **添付ファイル保存＆解凍**。<br>選択したメールの添付ファイルを所定フォルダに保存し、Zip/7zであればパスワードリストを用いて自動解凍を試みます。 |
| **modAutoFlag** | **メール自動フラグ**。<br>受信メールの本文を正規表現で解析し、「重要」「緊急」などのキーワードがあれば自動的にフラグを立てます。 |
| **modMailToPDF** | **メールのPDF化**。<br>選択したメールをPDFプリンタ経由で素早くPDF化します。処理後は元のプリンタへ自動的に復元されます。 |
| **modLaunchOWA** | **Web版Outlook起動**。<br>現在使用しているアカウントで OWA (Outlook on the Web) を素早くブラウザで開きます。 |
| **modMailOpen** | **メール開封制御**。<br>マウスクリック等による誤ったポップアップを防ぎ、ショートカットキーでのみメールを開封できるように制御します。 |
| **modLogger** | **共通ログ管理**。<br>全モジュールの動作ログを記録・管理します。自動アーカイブ機能付き。 |
## 必要要件
* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* [7-Zip](https://7-zip.opensource.jp/) (modSendController, modMailSevenZip で使用)
## インストール方法
1. **ソースコードのインポート**
    * Outlookで `Alt + F11` を押して VBA エディタを開きます。
    * `src/` フォルダ内の `.bas` ファイルをすべてインポートします。
    * **必須**: `modLogger.bas`
    * **選択**: 使用したい機能のモジュール（例: `modSendController.bas`）
2. **設定ファイルの配置**
    * エクスプローラーで `%APPDATA%` を開き、`OutlookVBA` という名前のフォルダを作成します。
    * パス例: `C:\Users\ユーザー名\AppData\Roaming\OutlookVBA\`
    * `configs/` フォルダ内のサンプルファイルを参考に、以下のファイルを作成・配置してください（文字コードは **UTF-8**）。

        | ファイル名 | 用途 | 元ファイル(参考) |
        | --- | --- | --- |
        | **config.ini** | 全ツールの統合設定 | `configs/config.sample.ini` |
        | **SevenZipPasswords.txt** | 7-Zip解凍用パスワードリスト | `configs/SevenZipPasswords.sample.txt` |

3. **マクロの有効化**
    * `src/ThisOutlookSession.cls` の内容を参考に、Outlookの `ThisOutlookSession` モジュールにコードを記述します。
    * これにより、メール受信時や送信時に自動的にツールが実行されるようになります。
## ディレクトリ構成
* `src/`: VBAソースコード
* `configs/`: 設定ファイルのサンプル (ini, txt)
* `docs/`: 詳細ドキュメントとコーディング規約
## ライセンス
MIT License