Attribute VB_Name = "modLogger"
Option Explicit

' ==============================================================================
' Module: modLogger
' Description: 共通ログ管理モジュール（ミリ秒対応・UTF-8・自動ZIPアーカイブ機能付）
' Dependencies: Scripting.FileSystemObject, ADODB.Stream, WScript.Shell
' Configuration: %APPDATA%\OutlookVBA\config.ini ([Logger], [General])
' ==============================================================================

' --- グローバル変数 ---
Private g_RunId As String
Private g_LogDir As String
Private g_ArchiveDays As Integer
Private g_7zPath As String
Private m_FSO As Object

Private Function GetFSO() As Object
    If m_FSO Is Nothing Then Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set GetFSO = m_FSO
End Function

' ==============================================================================
' [Public] 公開インターフェース
' ==============================================================================

Public Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

Public Sub Log(ByVal moduleName As String, ByVal msg As String)
    If g_LogDir = "" Then LoadLoggerConfig
    If Len(g_RunId) = 0 Then g_RunId = "NoID"
    
    Dim logMsg As String
    logMsg = GetTimestampWithMs() & " [" & g_RunId & "] [" & moduleName & "] " & msg
    
    Debug.Print logMsg
    AppendLogToFile logMsg
End Sub

' 他モジュールから7-Zipパスを取得するための公開関数
Public Function GetSevenZipPath() As String
    If g_LogDir = "" Then LoadLoggerConfig
    GetSevenZipPath = g_7zPath
End Function

' ==============================================================================
' [Config] 設定読み込み (modConfig 連携)
' ==============================================================================

Private Sub LoadLoggerConfig()
    ' デフォルト値
    g_LogDir = Environ$("APPDATA") & "\OutlookVBA\logs"
    g_ArchiveDays = 7
    g_7zPath = "C:\Program Files\7-Zip\7z.exe"
    
    Dim val As String
    val = modConfig.GetConfigValue("Logger", "LogDir", g_LogDir)
    g_LogDir = Replace(val, "%APPDATA%", Environ$("APPDATA"), 1, -1, vbTextCompare)
    
    val = modConfig.GetConfigValue("Logger", "ArchiveDays", "7")
    If IsNumeric(val) Then g_ArchiveDays = CInt(val)
    
    g_7zPath = modConfig.GetConfigValue("General", "SevenZipPath", g_7zPath)
    
    ArchiveOldLogs
End Sub

' ==============================================================================
' [Logic] ファイル操作・アーカイブ
' ==============================================================================

Private Sub AppendLogToFile(ByVal text As String)
    On Error Resume Next
    Dim fso As Object: Set fso = GetFSO()
    If Not fso.FolderExists(g_LogDir) Then CreateFolderRecursive fso, g_LogDir
    
    Dim filePath As String: filePath = g_LogDir & "\" & Format(Now, "yyyy-mm-dd") & ".log"
    Dim fileExists As Boolean: fileExists = fso.FileExists(filePath)
    
    ' 1行分の UTF-8 バイト列のみを生成
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8": stm.Open
    stm.WriteText text & vbCrLf
    
    ' 既存ファイルがある場合は BOM (3バイト) をスキップして末尾追記
    If fileExists Then
        stm.Position = 3
    Else
        stm.Position = 0
    End If
    stm.Type = 1 ' adTypeBinary
    
    Dim utf8Bytes() As Byte
    utf8Bytes = stm.Read
    stm.Close: Set stm = Nothing
    
    ' ネイティブバイナリモードでファイル末尾に直接追記 (O(1)、全ファイル再読み込みなし)
    Dim fn As Integer: fn = FreeFile
    Open filePath For Binary Access Write As #fn
    Seek #fn, LOF(fn) + 1
    Put #fn, , utf8Bytes
    Close #fn
    On Error GoTo 0
End Sub

Private Sub ArchiveOldLogs()
    On Error Resume Next
    Dim fso As Object: Set fso = GetFSO()
    If Dir$(g_7zPath) = "" Or Not fso.FolderExists(g_LogDir) Then Exit Sub
    
    Dim f As Object, targetDate As Date
    targetDate = DateAdd("d", -g_ArchiveDays, Date)
    
    Dim sh As Object: Set sh = Nothing
    
    For Each f In fso.GetFolder(g_LogDir).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "log" And f.DateLastModified < targetDate Then
            Dim zipPath As String: zipPath = f.Path & ".zip"
            If Not fso.FileExists(zipPath) Then
                If sh Is Nothing Then Set sh = CreateObject("WScript.Shell")
                sh.Run """" & g_7zPath & """ a """ & zipPath & """ """ & f.Path & """ -sdel", 0, True
            End If
        End If
    Next f
    Set sh = Nothing
    On Error GoTo 0
End Sub

Private Sub CreateFolderRecursive(ByVal fso As Object, ByVal path As String)
    Dim p As String: p = fso.GetParentFolderName(path)
    If Not fso.FolderExists(p) Then CreateFolderRecursive fso, p
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function GetTimestampWithMs() As String
    Dim t As Double: t = Timer
    Dim ms As Long: ms = CLng((t - Fix(t)) * 1000)
    If ms > 999 Then ms = 999
    GetTimestampWithMs = Format(Now, "yyyy/mm/dd hh:nn:ss") & "." & Format(ms, "000")
End Function
