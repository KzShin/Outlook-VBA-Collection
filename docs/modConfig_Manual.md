# Outlook VBA Common Config Loader (modConfig)

これは、Outlook VBA開発のための**共通設定管理モジュール**です。  
設定ファイル `%APPDATA%\OutlookVBA\config.ini`（UTF-8）の読み込みとメモリキャッシュを一元管理し、各マクロが高速かつ安全に設定値へアクセスできるようにします。

---

## 特徴

* **メモリキャッシュ機構**:
  * 初回アクセス時に `config.ini` を一度だけ読み込み、セクション単位で `Scripting.Dictionary` にキャッシュします。
  * 各マクロが設定値を参照するたびにディスク読み込み（I/O）を発生させず、高速に動作します。
* **大文字・小文字を区別しない柔軟な検索**:
  * セクション名およびキー名は大文字・小文字を区別しない（`vbTextCompare`）ため、記述の揺れによる設定取得失敗を防ぎます。
* **UTF-8対応**:
  * 日本語を含む設定値（宛先名、キーワードリスト、フォルダパスなど）を文字化けせずに読み込みます。
* **柔軟なコメント記法**:
  * 行頭の `#`、`;`、`//` および空行を自動的にスキップします。

---

## 必要要件

* Windows 10 / 11
* Microsoft Outlook (Classic Desktop)
* 参照設定: 標準で動作（COMオブジェクト `Scripting.FileSystemObject`, `Scripting.Dictionary`, `ADODB.Stream` をレイトバインドで使用）

---

## インストール

1. Outlookを起動し、`Alt + F11` でVBAエディタを開きます。
2. **モジュールのインポート**:
   * `File` > `Import File` から **`src/modConfig.bas`** をインポートします。
   * ※各機能モジュール（`modSendController`, `modMailSevenZip`, `modLogger` 等）を使用する際の**必須モジュール**となります。

---

## 設定ファイル (config.ini) の配置

1. エクスプローラーのアドレスバーに `%APPDATA%\OutlookVBA\` を開きます。
2. `configs/config.sample.ini` をコピーし、`config.ini` という名前で配置します。
3. 文字コードは **UTF-8** で保存してください。

```ini
[General]
SevenZipPath=C:\Program Files\7-Zip\7z.exe

[Logger]
LogDir=%APPDATA%\OutlookVBA\logs
ArchiveDays=7

[SendController]
HolidayList=12-29,12-30,12-31,01-01,01-02,01-03
WorkStartTime=08:00
WorkEndTime=18:00
```

---

## 公開関数 (API) リファレンス

### 1. `GetConfigValue`
指定したセクション名とキー名に対応する値を取得します。

```vb
Public Function GetConfigValue(ByVal sectionName As String, _
                               ByVal keyName As String, _
                               Optional ByVal defaultValue As String = "") As String
```

* **引数**:
  * `sectionName`: INIファイルのセクション名（例: `"General"`, `"Logger"`）
  * `keyName`: キー名（例: `"SevenZipPath"`, `"LogDir"`）
  * `defaultValue` (省略可能): キーが存在しない場合に返すデフォルト文字列
* **戻り値**: 設定値（文字列）。キーが存在しない場合は `defaultValue` を返します。
* **使用例**:
  ```vb
  Dim sevenZip As String
  sevenZip = modConfig.GetConfigValue("General", "SevenZipPath", "C:\Program Files\7-Zip\7z.exe")
  ```

---

### 2. `GetSection`
指定したセクションに含まれるすべてのキーと値を `Scripting.Dictionary` として取得します。

```vb
Public Function GetSection(ByVal sectionName As String) As Object
```

* **引数**:
  * `sectionName`: セクション名（例: `"ForwardMail"`）
* **戻り値**: `Scripting.Dictionary` オブジェクト（Key -> Value）。セクションが存在しない場合は空のDictionaryが返ります（`Nothing` は返しません）。
* **使用例**:
  ```vb
  Dim dict As Object, k As Variant
  Set dict = modConfig.GetSection("ForwardMail")
  
  For Each k In dict.Keys
      Debug.Print "Key: " & k & " / Value: " & dict(k)
  Next k
  ```

---

### 3. `ReloadConfig`
保持しているメモリキャッシュを破棄し、ディスク上の `config.ini` を再読み込みします。

```vb
Public Sub ReloadConfig()
```

* **使用例**:
  * マクロ実行中に `config.ini` の内容を手動で書き換えた直後などに呼び出します。

---

### 4. `GetConfigPath`
設定ファイル `config.ini` の標準フルパスを取得します。

```vb
Public Function GetConfigPath() As String
' 戻り値: "%APPDATA%\OutlookVBA\config.ini" の実パス
```

---

### 5. `IsLoaded`
設定が既にメモリ上にロードされているかを確認します。

```vb
Public Function IsLoaded() As Boolean
```
