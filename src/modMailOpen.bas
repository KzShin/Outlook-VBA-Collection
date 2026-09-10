Attribute VB_Name = "modMailOpen"
Option Explicit

' ==============================================================================
' Module: modMailOpen
' Description: ショートカットキー操作によるメールのポップアップ開封を制御する
' Dependencies: Scripting.FileSystemObject, ADODB.Stream, modLogger
' Configuration: %APPDATA%\OutlookVBA\config.ini ([MailOpen] Section)
' ==============================================================================

' --- グローバル変数 (Module Level) ---
Private g_RunId As String

' ThisOutlookSessionから参照するため、このフラグのみ例外的にPublicとします
Public g_AllowOpen As Boolean 

' ==============================================================================
' [Private] ログ・ID管理ヘルパー
' ==============================================================================

Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

' 共通ロガーへの委譲
Private Sub Log(ByVal msg As String)
    modLogger.Log "MailOpen", msg
End Sub

' ==============================================================================
' [Public] 公開インターフェース
' ==============================================================================

' 機能の有効/無効を判定する（ThisOutlookSessionからも呼び出されます）
Public Function IsFeatureEnabled() As Boolean
    Dim val As String
    val = GetConfigValue("EnableShortcutOpen", "False") ' デフォルトは False(オフ)
    IsFeatureEnabled = (LCase$(val) = "true")
End Function

' ショートカットから呼び出されるマクロ
Public Sub OpenSelectedMail()
    On Error GoTo EH
    
    ' 機能がオフの場合はマクロを実行しない
    If Not IsFeatureEnabled() Then Exit Sub
    
    ' 1. 実行IDの生成とセット (yymmdd-hhnnss-OPEN)
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-OPEN"
    
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START 選択メール開封処理 ==="
    
    ' 2. エクスプローラーと選択アイテムの取得
    Dim exp As Outlook.Explorer
    Set exp = Application.ActiveExplorer
    
    If exp Is Nothing Then
        Log "ActiveExplorerが見つかりません。処理を中止します。"
        GoTo FIN
    End If
    
    If exp.Selection.Count = 0 Then
        Log "アイテムが選択されていません。処理を中止します。"
        GoTo FIN
    End If
    
    Dim objItem As Object
    Set objItem = exp.Selection(1)
    
    ' 3. メールアイテムの判定と開封処理
    If objItem.Class = olMail Then
        Dim mailItem As Outlook.MailItem
        Set mailItem = objItem
        
        Log "対象メールを開封します: " & mailItem.Subject
        
        ' 開封許可フラグを立ててからDisplayを呼び出す
        g_AllowOpen = True
        On Error Resume Next
        mailItem.Display
        g_AllowOpen = False
        On Error GoTo EH
    Else
        Log "選択アイテムはメールではありません (Class=" & objItem.Class & ")"
    End If

FIN:
    Log "=== END 選択メール開封処理 ==="
    modLogger.SetRunId "NoID"
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    ' エラー時も安全のためフラグを確実に下ろす
    g_AllowOpen = False 
    modLogger.SetRunId "NoID"
End Sub

' ==============================================================================
' [Config] 設定読み込み (modConfig 連携)
' ==============================================================================

Private Function GetConfigValue(ByVal key As String, Optional ByVal defaultVal As String = "") As String
    Dim val As String
    val = modConfig.GetConfigValue("MailOpen", key, "")
    If Len(val) > 0 Then
        GetConfigValue = val
    Else
        If Len(defaultVal) > 0 Then
            Log "Config Key Not Found: [" & key & "] -> Using Default: " & defaultVal
        End If
        GetConfigValue = defaultVal
    End If
End Function
