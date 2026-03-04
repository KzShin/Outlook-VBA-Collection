# Outlook VBA コーディング規約 (General Coding Standards)

## 1. 基本方針 (General Principles)

GitHub等での公開および複数環境での配布を前提とし、以下の原則を遵守する。

1. **Environment Independence（環境独立性）**:
    * 特定のユーザー名やドライブパス（例: `C:\Users\Taro\...`）をハードコーディングしない。
    * 必ず `Environ("APPDATA")` や `Environ("USERPROFILE")` を使用してパスを動的に生成する。


2. **No Secrets in Code（機密情報の排除）**:
    * メールアドレス、パスワード、APIキー、社内サーバー名などをコード内に直接記述しない。
    * すべて外部設定ファイル（`config.ini` 等）から読み込む設計とする。


3. **UTF-8 Standardization（UTF-8標準化）**:
    * 設定ファイルやログ、テキスト出力のエンコードは、Shift-JISではなく **UTF-8** を標準とする。


4. **Self-Contained Logic（自己完結性）**:
    * 標準モジュール単体で機能が完結するように設計し、外部依存（他モジュールへの依存）を最小限にする。
    * **例外**: ログ出力に関しては、共通モジュール `modLogger` への依存を許容・推奨する。


## 2. モジュール構成 (Module Structure)

標準モジュールは以下のセクション構成で統一する。冒頭のヘッダーコメントは必須とする。

```vb
Attribute VB_Name = "modModuleName"
Option Explicit

' ==============================================================================
' Module: modModuleName (例: modMyFeature)
' Description: モジュールの概要（何をするものか）
' Dependencies: modLogger, Scripting.FileSystemObject ... (依存ライブラリ)
' Configuration: %APPDATA%\OutlookVBA\config.ini ([MySectionName])
' ==============================================================================

' --- グローバル変数 (Module Level) ---
Private g_RunId As String  ' 実行単位の一意ID

' ==============================================================================
' [Public] 公開インターフェース・初期化
' ==============================================================================
' (外部から呼び出すSub/Function。ここでRunIDを生成し、modLoggerへ渡す)

' ==============================================================================
' [Main] メイン処理
' ==============================================================================
' (処理の核心部分)

' ==============================================================================
' [Config] 設定・データ管理
' ==============================================================================
' (外部ファイルの読み書き。UTF-8対応必須)

' ==============================================================================
' [Logic] ビジネスロジック・判定
' ==============================================================================
' (純粋な計算や判定処理)
```


## 3. 命名規則 (Naming Conventions)

GitHubでの可読性を考慮し、識別子は基本的に**英語**を使用する。（コメントは日本語で可）

| 対象 | 規則 | 例 | 備考 |
| --- | --- | --- | --- |
| **モジュール** | `mod` + PascalCase | `modTaskAutomator`, `modLogger` | 機能を表す名詞 |
| **プロシージャ** | PascalCase | `LoadConfig`, `ProcessEmail` | 動詞 + 名詞 |
| **変数 (ローカル)** | camelCase | `folderPath`, `targetMail` |  |
| **変数 (モジュール)** | `g_` + PascalCase | `g_RunId`, `g_ConfigCache` | Private変数でも識別しやすくする |
| **定数** | ALL_UPPER_CASE | `MAX_RETRY`, `DEFAULT_PATH` | アンダースコア区切り |
| **引数** | camelCase | `inputData`, `isRecursive` |  |


## 4. ログ出力の実装 (Logging Standard)

デバッグとトラブルシューティングのため、 **共通モジュール** `modLogger`  を利用した統一ログ出力を実装する。
各モジュールで独自のログファイル出力ロジックを持たず、すべて `modLogger` に委譲する。

### 4.1. 共通プロシージャの定義（各モジュール内）

各モジュールの `[Private]` セクションに以下を実装し、`modLogger` をラップする。

```vb
' 実行IDを保持
Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

' 共通Loggerへの委譲
Private Sub Log(ByVal msg As String)
    ' 第1引数に自モジュール名を渡す
    modLogger.Log "MyModuleName", msg
End Sub
```

### 4.2. 利用ルール（エントリーポイント）

Public プロシージャの冒頭で ID を生成し、自モジュールと `modLogger` の両方にセットする。

```vb
Public Sub MainProcess(ByVal Item As Object)
    On Error GoTo EH
    
    ' 【必須】処理開始時にIDを生成
    ' フォーマット: yymmdd-hhnnss-識別子 (例: 260216-120000-SAVE)
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-MYPROC"
    
    ' 自モジュールと共通ロガーの両方にIDをセット
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START 処理名 ==="
    ' ... 処理本体 ...
    Log "=== END 処理名 ==="
    Exit Sub
EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
End Sub
```

* **出力形式**: `yyyy/mm/dd hh:nn:ss.ms [RunID] [ModuleName] メッセージ`
* **ログファイル**: `%APPDATA%\OutlookVBA\logs\yyyy-mm-dd.log` (UTF-8, 自動ZIP圧縮)


## 5. 設定ファイルの管理 (Configuration)

設定値はハードコーディングせず、必ず外部ファイルから読み込む。
管理を容易にするため、**1つの統合設定ファイル (`config.ini`)** を使用し、モジュールごとに**セクション**を分けて管理する。

### 5.1. 保存場所と形式

* **パス**: `%APPDATA%\OutlookVBA\config.ini`
* **形式**: INI形式。セクション `[SectionName]` で区切る。
* **文字コード**: **UTF-8**。
* **コメント**: 行頭が `#` の行はコメントとして扱い、読み込み時に無視する仕様とする。

### 5.2. 読み込みの実装

`ADODB.Stream` を使用してUTF-8で読み込み、**「現在のセクション」を判定して自モジュールの設定のみを取得する** ロジックを実装する。

**推奨読み込みコード（セクション対応版）:**

```vb
Dim stm As Object
Set stm = CreateObject("ADODB.Stream")
' UTF-8で全読み込み
With stm
    .Type = 2          ' adTypeText
    .Charset = "UTF-8"
    .Open
    .LoadFromFile configPath
End With
Dim allText As String
allText = stm.ReadText(-1)
stm.Close

' 行ごとに処理
Dim lines() As String
Dim i As Long, lineText As String
Dim currentSection As String ' 現在のセクションを保持

lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)

For i = LBound(lines) To UBound(lines)
    lineText = Trim$(lines(i))
    
    ' 空行とコメント(#)をスキップ
    If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
        
        ' セクション開始の判定 [SectionName]
        If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
            currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
        
        ' 対象セクション（例: myfeature）の場合のみ読み込み
        ElseIf currentSection = "myfeature" Then
            ' Key=Value 解析
            ' ...
        End If
    End If
Next i
```


## 6. ライブラリとオブジェクト参照 (References)

配布時のエラーを防ぐため、可能な限り「遅延バインディング (Late Binding)」を使用する。

| ライブラリ | 推奨される記述 | 理由 |
| --- | --- | --- |
| **FileSystemObject** | `CreateObject("Scripting.FileSystemObject")` | 参照設定不要で動作させるため |
| **RegExp** | `CreateObject("VBScript.RegExp")` | 同上 |
| **ADODB** | `CreateObject("ADODB.Stream")` | 同上 |
| **WScript.Shell** | `CreateObject("WScript.Shell")` | 外部コマンド実行用 |


## 7. エラー処理 (Error Handling)

1. **エントリーポイント**: Publicプロシージャには必ず `On Error GoTo` を含め、予期せぬ終了を防ぐ。
2. **部分的な無視**: `On Error Resume Next` は、ファイル存在確認やループ内のスキップ処理など、限定的な範囲でのみ使用し、直後に `On Error GoTo 0` で解除する。
3. **ユーザー通知**: エラー発生時は `Log` に出力し、必要であればメッセージボックス等でユーザーに通知する。


## 8. GitHub公開時のチェックリスト

リポジトリへの Push 前に以下を確認する。

1. [ ] **個人情報の削除**: コード内のメールアドレスやパスワードがダミー化されているか。
2. [ ] **サンプル設定ファイル**: `config.ini` を直接含めず、全モジュールの設定例を網羅した `configs/config.sample.ini` を同梱しているか。
3. [ ] **機密ファイルの除外**: `.gitignore` に `config.ini` や `SevenZipPasswords.txt` が含まれているか。
4. [ ] **依存関係の記述**: ヘッダーコメントに `modLogger` や外部ツール（7-Zip等）への依存、および使用する `[SectionName]` が記載されているか。