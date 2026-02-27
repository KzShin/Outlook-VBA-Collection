# Outlook Mail to PDF Converter (modMailToPDF)
これは、OutlookのVBAマクロです。選択したメールを「Microsoft Print to PDF」などのPDF仮想プリンタを使用して、素早くPDFファイルとして保存します。
処理中は一時的にデフォルトプリンタをPDFプリンタへ切り替え、PDF出力が完了すると自動的に元のプリンタ（通常使用している物理プリンタなど）へ設定を復元します。
## 特徴
* **ワンクリックPDF化**: リボンやクイックアクセスツールバーに登録することで、選択中のメールを即座にPDF化できます。
* **元プリンタの自動復元機能**: PDF出力後、元のプリンタに自動で戻します。設定を `Auto` にしておけば、ノートPCの持ち歩きなどで通常使うプリンタが変わる環境でも、マクロ実行時点のデフォルトプリンタを自動判別して復元します。
* **プリンタのカスタマイズ**: 使用するPDFプリンタや復元先のプリンタ名を設定ファイルから自由に変更可能です。
* **共通ログ機能**: 共通モジュール `modLogger` を使用し、実行履歴やエラー（プリンタ設定失敗など）を記録します。
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
        2. `src/modMailToPDF.bas` (メイン機能モジュール)
4. **ThisOutlookSessionの設定**:
    * VBAエディタ左側の `Project1` > `Microsoft Outlook Objects` > `ThisOutlookSession` をダブルクリックします。
    * 同梱の `src/ThisOutlookSession.cls` の中身をコピーし、貼り付けます。
    * または、以下のプロシージャを追加してください。
        ```vb
        ' 選択したメールをPDFとして印刷するマクロ
        Public Sub 選択メールをPDF化()
            ' modMailToPDF内でRunID生成・ログ出力まで完結するため、直接呼び出します
            modMailToPDF.PrintMailToPDF
        End Sub
        ```
## 設定 (Configuration)
本ツールは、共通設定ファイル `%APPDATA%\OutlookVBA\config.ini` の **`[MailToPDF]`** セクションを使用します。
1. **設定ファイルの準備**:
    * エクスプローラーで `%APPDATA%\OutlookVBA\` を開きます。
    * `config.ini` をテキストエディタで開きます（ファイルがない場合は `configs/config.sample.ini` をコピーして作成）。
2. **設定の編集**:
    * 以下のセクションを追記・編集します。
    * **重要**: 保存時の文字コードは必ず **UTF-8 (BOMなし推奨)** にしてください。
        ```ini
        [MailToPDF]
        # PDF出力用プリンタ名
        PdfPrinterName=Microsoft Print to PDF
        # 出力後に戻す元のプリンタ名
        # （空欄または Auto にすると、実行時点のデフォルトプリンタを自動で判別して元に戻します）
        PhysicalPrinterName=Auto
        ```
## 使い方
このマクロは、ボタンに割り当てて使用することを推奨します。
1. **ボタンの登録**:
    * Outlookのウインドウ上部で右クリックし、「リボンのユーザー設定」または「クイックアクセスツールバーのユーザー設定」を開きます。
    * 「コマンドの選択」で「マクロ」を選びます。
    * `Project1.ThisOutlookSession.選択メールをPDF化` を選択し、右側のリストに追加します。
    * 必要に応じて「名前の変更」でアイコンや表示名（「PDF保存」など）を変更します。
2. **実行**:
    * PDF化したいメールを選択状態にします。
    * 追加したボタンをクリックすると、PDFの保存先とファイル名を指定するダイアログが表示されます。
    * 任意の場所に保存すると、自動的に元のプリンタ設定に戻ります。
## ログ
本ツールは `modLogger` を使用して動作ログを記録します。
* **ログ保存場所**: `%APPDATA%\OutlookVBA\logs\`
* **ファイル名**: `yyyy-mm-dd.log` (日付ごとのローテーション)
* **ログ出力例:**
    ```text
    2026/02/16 10:00:00.123 [260216-100000-TOPDF] [MailToPDF] === START メールPDF化処理 ===
    2026/02/16 10:00:00.130 [260216-100000-TOPDF] [MailToPDF] 対象メール: プロジェクト進捗報告
    2026/02/16 10:00:00.140 [260216-100000-TOPDF] [MailToPDF] 元プリンタを自動取得しました: Canon XXXX Series
    2026/02/16 10:00:00.150 [260216-100000-TOPDF] [MailToPDF] デフォルトプリンタを変更: Microsoft Print to PDF
    2026/02/16 10:00:00.160 [260216-100000-TOPDF] [MailToPDF] 印刷開始 (PrintOut)
    2026/02/16 10:00:05.000 [260216-100000-TOPDF] [MailToPDF] 印刷コマンド送信完了
    2026/02/16 10:00:05.010 [260216-100000-TOPDF] [MailToPDF] デフォルトプリンタを復元: Canon XXXX Series
    2026/02/16 10:00:05.020 [260216-100000-TOPDF] [MailToPDF] === END メールPDF化処理 ===
    ```
## ライセンス
MIT License
