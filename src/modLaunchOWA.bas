Attribute VB_Name = "modLaunchOWA"
Option Explicit

' ==============================================================================
' Module: modLaunchOWA
' Description: WEB版Outlook (OWA) をブラウザで開く（検索機能なし）
' Dependencies: Scripting.FileSystemObject, WScript.Shell, ADODB.Stream, modLogger
' Configuration: %APPDATA%\OutlookVBA\config.ini ([LaunchOWA] Section)
' ==============================================================================

' --- ThisOutlookSessionでの呼び出し例 ---
' Public Sub CustomButton_Click()
'     ' リボンやクイックアクセスツールバーのボタンに割り当て
'     modLaunchOWA.LaunchOWA
' End Sub

' --- グローバル変数 (Module Level) ---
Private g_RunId As String
Private Const DEFAULT_URL As String = "https://outlook.office.com/mail/"

' ==============================================================================
' [Public] 公開インターフェース
' ==============================================================================

' ユーザーがボタンから実行するエントリポイント
Public Sub LaunchOWA()
    On Error GoTo EH
    
    ' 実行ID生成 (yymmdd-hhnnss)
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss")
    
    ' 自モジュールと共通ロガーの両方にIDを設定
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START OWA Launch ==="
    
    ' メイン処理
    OpenOutlookWeb
    
    Log "=== END OWA Launch ==="
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "WEB版Outlookの起動中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical, "Outlook OWA Launcher"
End Sub

' --- 内部ヘルパー ---

Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

' 共通ロガーへの委譲
Private Sub Log(ByVal msg As String)
    modLogger.Log "LaunchOWA", msg
End Sub

' ==============================================================================
' [Main] メイン処理
' ==============================================================================

Private Sub OpenOutlookWeb()
    Dim targetUrl As String
    
    ' 1. 設定ファイルからURLを取得
    targetUrl = GetConfigValue("BaseUrl")
    
    ' 2. 設定がない場合はデフォルトを使用
    If Len(targetUrl) = 0 Then
        targetUrl = DEFAULT_URL
        Log "Config not found or BaseUrl empty. Using default URL: " & targetUrl
    Else
        Log "Loaded URL from config: " & targetUrl
    End If
    
    ' 3. ブラウザ起動
    Log "Opening URL..."
    OpenUrl targetUrl
End Sub

' ==============================================================================
' [Config] 設定管理
' ==============================================================================

Private Function GetConfigValue(ByVal targetKey As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim configPath As String
    ' 変更: 共通設定ファイルのパス
    configPath = Environ("APPDATA") & "\OutlookVBA\config.ini"
    
    If Not fso.FileExists(configPath) Then
        Log "Config file missing: " & configPath
        Exit Function
    End If
    
    ' ADODB.StreamによるUTF-8読み込み
    On Error Resume Next
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    
    With stm
        .Type = 2          ' adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile configPath
    End With
    
    Dim allText As String
    allText = stm.ReadText(-1)
    stm.Close
    On Error GoTo 0
    
    ' 行ごとの解析
    Dim lines() As String
    lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    
    Dim i As Long, lineText As String
    Dim separatorPos As Long
    Dim currentSection As String
    Dim key As String, value As String
    
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        
        ' コメント(#)と空行スキップ
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
            
            ' セクション判定 [SectionName]
            If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
                
            ' [LaunchOWA] セクションのみ処理
            ElseIf currentSection = "launchowa" Then
                separatorPos = InStr(lineText, "=")
                
                If separatorPos > 0 Then
                    key = Trim$(Left$(lineText, separatorPos - 1))
                    value = Trim$(Mid$(lineText, separatorPos + 1))
                    
                    If LCase$(key) = LCase$(targetKey) Then
                        GetConfigValue = value
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' ==============================================================================
' [Logic] ブラウザ起動
' ==============================================================================

Private Sub OpenUrl(ByVal url As String)
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run url
    Set wsh = Nothing
End Sub
