# Outlook VBA Common Logger (modLogger)

これは、Outlook VBA開発のための**共通ログ管理モジュール**です。
複数のVBAマクロから利用されることを想定しており、イミディエイトウィンドウへの出力とログファイルへの保存（UTF-8）を一元管理します。また、古くなったログファイルの自動圧縮（Zip化）機能も備えています。

## 特徴

* **一元管理**: `Debug.Print`（イミディエイトウィンドウ）とファイル出力を同時に行います。
* **高精度タイムスタンプ**: ミリ秒単位（`yyyy/mm/dd hh:nn:ss.ms`）でのログ記録に対応しており、処理時間の計測に役立ちます。
* **UTF-8対応**: `ADODB.Stream` を使用しており、日本語を含むログも文字化けせずに保存します。
* **自動アーカイブ**: 指定日数を経過した古いログファイルを、**7-Zip** を使用して自動的にZIP圧縮し、元ファイルを削除します。
* **環境独立性**: ログの保存場所や設定ファイルは `%APPDATA%` を基準とするため、ユーザー環境に依存しません。

## 必要要件

* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* [7-Zip](https://7-zip.opensource.jp/)
* ログの自動圧縮機能を使用するために必要です。
* デフォルトパス: `C:\Program Files\7-Zip\7z.exe` (設定ファイルで変更可能)



## インストール

1. このリポジトリのファイルをダウンロードします。
2. Outlookを起動し、`Alt + F11` でVBAエディタを開きます。
3. **モジュールのインポート**:
    * `File` > `Import File` から `src/modLogger.bas` をインポートします。


4. **設定ファイルの配置**:
    * 下記「設定 (Configuration)」セクションの手順に従って初期設定を行います。



## 設定 (Configuration)

本モジュールは、共通設定ファイル `%APPDATA%\OutlookVBA\config.ini` の **`[Logger]`** および **`[General]`** セクションを使用します。

1. **フォルダの作成**:
    * エクスプローラーのアドレスバーに `%APPDATA%` と入力して移動します。
    * `OutlookVBA` という名前のフォルダを新規作成します。
    * パス例: `C:\Users\ユーザー名\AppData\Roaming\OutlookVBA\`


2. **設定ファイルの配置**:
    * `configs/config.sample.ini` を参考に `config.ini` を作成し、配置します。
    * ※文字コードは必ず **UTF-8 (BOMなし推奨)** で保存してください。


3. **設定の編集**:
    * `config.ini` をテキストエディタで開き、環境に合わせて修正します。


        ```ini
        [General]
        # [共通] 7-Zipの実行ファイルパス
        # modLoggerの自動アーカイブ機能で使用されます。
        SevenZipPath=C:\Program Files\7-Zip\7z.exe

        [Logger]
        # [共通] ログの保存設定
        # 保存先フォルダ (%APPDATA% 変数使用可)
        LogDir=%APPDATA%\OutlookVBA\logs

        # 何日経過したログをZIP圧縮するか
        ArchiveDays=7

        ```



## 使い方 (開発者向け)

他の標準モジュールや `ThisOutlookSession` から、以下のように呼び出して使用します。

### 1. 実行IDの設定

処理の開始時（エントリーポイント）で `SetRunId` を呼び出し、一連の処理IDを設定します。これにより、並行して走る複数の処理をログ上で区別できます。

```vb
' 推奨フォーマット: yymmdd-hhnnss-識別子
Dim runId As String
runId = Format(Now, "yymmdd-hhnnss") & "-MYPROC"

' IDをセット
modLogger.SetRunId runId

```

### 2. ログ出力

`Log` メソッドを呼び出します。第1引数に呼び出し元のモジュール名、第2引数にメッセージを指定します。

```vb
modLogger.Log "MyModule", "処理を開始します"
' ...
modLogger.Log "MyModule", "件数: " & count

```

## ログ出力仕様

* **保存先**: `%APPDATA%\OutlookVBA\logs\yyyy-mm-dd.log` (設定により変更可)
* **フォーマット**: `yyyy/mm/dd hh:nn:ss.ms [RunID] [ModuleName] Message`

    **出力例:**

    ```text
    2026/02/16 10:00:00.123 [260216-100000-SAVE] [ThisOutlookSession] === START Process ===
    2026/02/16 10:00:00.150 [260216-100000-SAVE] [MailSevenZip] 対象メール: Project-A Report
    2026/02/16 10:00:01.005 [260216-100000-SAVE] [MailSevenZip] 展開成功: C:\Users\User\Downloads\...
    2026/02/16 10:00:01.020 [260216-100000-SAVE] [ThisOutlookSession] === END Process ===

    ```

## ライセンス

MIT License