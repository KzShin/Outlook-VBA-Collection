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
' [Config] 設定読み込み (統合版)
' ==============================================================================

Private Sub LoadLoggerConfig()
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim configPath As String
    ' 変更: 統合設定ファイルを参照
    configPath = Environ$("APPDATA") & "\OutlookVBA\config.ini"
    
    ' デフォルト値
    g_LogDir = Environ$("APPDATA") & "\OutlookVBA\logs"
    g_ArchiveDays = 7
    g_7zPath = "C:\Program Files\7-Zip\7z.exe"
    
    If fso.FileExists(configPath) Then
        On Error Resume Next
        Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
        With stm
            .Type = 2: .Charset = "UTF-8": .Open: .LoadFromFile configPath
        End With
        Dim allText As String: allText = stm.ReadText(-1)
        stm.Close
        On Error GoTo 0
        
        Dim lines() As String: lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
        Dim i As Long, lineText As String, parts() As String
        Dim currentSection As String
        
        For i = LBound(lines) To UBound(lines)
            lineText = Trim$(lines(i))
            If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
                ' セクション判定
                If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                    currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
                
                ' [Logger] または [General] セクションを読み込む
                ElseIf currentSection = "logger" Or currentSection = "general" Then
                    parts = Split(lineText, "=", 2)
                    If UBound(parts) = 1 Then
                        Select Case Trim$(parts(0))
                            Case "LogDir": g_LogDir = Replace(Trim$(parts(1)), "%APPDATA%", Environ$("APPDATA"))
                            Case "ArchiveDays": g_ArchiveDays = CInt(Trim$(parts(1)))
                            Case "SevenZipPath": g_7zPath = Trim$(parts(1))
                        End Select
                    End If
                End If
            End If
        Next i
    End If
    
    ArchiveOldLogs
End Sub

' ==============================================================================
' [Logic] ファイル操作・アーカイブ
' ==============================================================================

Private Sub AppendLogToFile(ByVal text As String)
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(g_LogDir) Then CreateFolderRecursive fso, g_LogDir
    
    Dim filePath As String: filePath = g_LogDir & "\" & Format(Now, "yyyy-mm-dd") & ".log"
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8": stm.Open
    
    If fso.FileExists(filePath) Then
        stm.LoadFromFile filePath: stm.Position = stm.Size
    End If
    stm.WriteText text & vbCrLf: stm.SaveToFile filePath, 2: stm.Close
    On Error GoTo 0
End Sub

Private Sub ArchiveOldLogs()
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Dir$(g_7zPath) = "" Or Not fso.FolderExists(g_LogDir) Then Exit Sub
    
    Dim f As Object, targetDate As Date
    targetDate = DateAdd("d", -g_ArchiveDays, Date)
    
    For Each f In fso.GetFolder(g_LogDir).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "log" And f.DateLastModified < targetDate Then
            Dim zipPath As String: zipPath = f.Path & ".zip"
            If Not fso.FileExists(zipPath) Then
                CreateObject("WScript.Shell").Run """" & g_7zPath & """ a """ & zipPath & """ """ & f.Path & """ -sdel", 0, True
            End If
        End If
    Next f
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
