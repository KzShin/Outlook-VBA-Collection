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
Private g_ConfigCache As Object     ' 設定キャッシュ (Dictionary)
Private g_IsConfigLoaded As Boolean ' 設定読み込み済みフラグ

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
        mailItem.Display
        g_AllowOpen = False
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
' [Config] 設定読み込み (統合版 config.ini 対応)
' ==============================================================================

Private Sub LoadConfig()
    On Error GoTo EH
    Set g_ConfigCache = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim configPath As String
    configPath = Environ$("APPDATA") & "\OutlookVBA\config.ini"
    
    If Not fso.FileExists(configPath) Then
        Log "Config file not found: " & configPath
        g_IsConfigLoaded = True
        Exit Sub
    End If
    
    ' ADODB.StreamによるUTF-8読み込み
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile configPath
    End With
    
    Dim allText As String: allText = stm.ReadText(-1)
    stm.Close
    
    ' 解析処理（セクション対応）
    Dim lines() As String: lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    Dim i As Long, lineText As String, eqPos As Long
    Dim key As String, val As String
    Dim currentSection As String
    
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        ' コメント(#)と空行をスキップ
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
            
            ' セクション判定 [SectionName]
            If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
            
            ' [MailOpen] セクションのみ読み込む
            ElseIf currentSection = "mailopen" Then
                eqPos = InStr(lineText, "=")
                If eqPos > 1 Then
                    key = Trim$(Left$(lineText, eqPos - 1))
                    val = Trim$(Mid$(lineText, eqPos + 1))
                    g_ConfigCache(key) = val
                End If
            End If
        End If
    Next i
    
    g_IsConfigLoaded = True
    Log "Config Loaded Successfully."
    Exit Sub
EH:
    Log "Config Load Error: " & Err.Description
    g_IsConfigLoaded = True
End Sub

Private Function GetConfigValue(ByVal key As String, Optional ByVal defaultVal As String = "") As String
    If Not g_IsConfigLoaded Then LoadConfig
    
    If g_ConfigCache.Exists(key) Then
        GetConfigValue = g_ConfigCache(key)
    Else
        Log "Config Key Not Found: [" & key & "] -> Using Default: " & defaultVal
        GetConfigValue = defaultVal
    End If
End Function
