Attribute VB_Name = "modSendController"
Option Explicit

' ==============================================================================
' Module: modSendController
' Description: メール送信制御（添付確認、Zip暗号化チェック、送信予約）
' Dependencies: Scripting.FileSystemObject, WScript.Shell, ADODB.Stream, modLogger
' Configuration: %APPDATA%\OutlookVBA\config.ini ([SendController] Section)
' ==============================================================================

' --- ThisOutlookSessionでの呼び出し例 ---
' Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'     modSendController.Execute Item, Cancel
' End Sub

' --- 定数 (Constants) ---
Private Const DEFAULT_HOLIDAYS As String = "12-29,12-30,12-31,01-01,01-02,01-03"

' --- グローバル変数 (Module Level) ---
Private g_RunId As String           ' 実行ログ用ID

' ==============================================================================
' [Public] 公開インターフェース
' ==============================================================================

' 送信制御の実行
Public Sub Execute(ByVal Item As Object, ByRef Cancel As Boolean)
    On Error GoTo EH
    
    ' 1. 初期化 (RunID生成: yymmdd-hhnnss-SEND)
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-SEND"
    
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START SendController ==="
    
    ' 2. 対象確認 (MailItem以外は除外)
    If Not TypeOf Item Is Outlook.MailItem Then
        Log "Target is not MailItem. Skip."
        Exit Sub
    End If
    
    Dim m As Outlook.MailItem
    Set m = Item
    Log "Subject=" & SafeStr(m.Subject) & " / Attachments=" & m.Attachments.Count
    
    ' --- フロー実行 ---
    
    ' Step 1: 添付忘れ確認
    Dim allowNoAttachment As Boolean
    If Not CheckAttachmentMention(m, allowNoAttachment) Then
        Cancel = True
        GoTo FIN
    End If
    
    ' Step 2: Zip/7z パスワード確認
    If Not allowNoAttachment And m.Attachments.Count > 0 Then
        If Not CheckZipPassword(m) Then
            Cancel = True
            GoTo FIN
        End If
    End If
    
    ' Step 3: 送信時刻制御
    If Not CheckSendTime(m) Then
        Cancel = True
        GoTo FIN
    End If

FIN:
    Log "=== END SendController / Cancel=" & CStr(Cancel) & " ==="
    SetRunId "NoID"
    modLogger.SetRunId "NoID"
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "送信処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    Cancel = True
    ' エラー時は念のため下書きへ退避
    SaveToDraftsSafe Item
    SetRunId "NoID"
    modLogger.SetRunId "NoID"
End Sub

' ==============================================================================
' [Private] ログ・ID管理ヘルパー
' ==============================================================================

Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

' 共通ロガーへの委譲
Private Sub Log(ByVal msg As String)
    modLogger.Log "SendController", msg
End Sub

Private Function FormatTime(ByVal startTime As Double) As String
    Dim elapsed As Double
    Dim t As Double: t = Timer
    If t >= startTime Then
        elapsed = t - startTime
    Else
        elapsed = (86400# - startTime) + t
    End If
    FormatTime = CStr(CLng(elapsed * 1000))
End Function

Private Function SafeStr(ByVal s As String) As String
    SafeStr = Replace(Replace(s, vbCrLf, " "), vbCr, " ")
End Function

' ==============================================================================
' [Logic 1] 添付ファイル確認ロジック
' ==============================================================================
Private Function CheckAttachmentMention(ByRef m As Outlook.MailItem, ByRef allowNoAttachment As Boolean) As Boolean
    CheckAttachmentMention = True ' デフォルト継続
    allowNoAttachment = False
    
    Dim t0 As Double: t0 = Timer
    
    ' 「添付」という文字があるか
    Dim hasWord As Boolean
    hasWord = (InStr(1, m.Body, "添付", vbTextCompare) > 0)
    
    If hasWord And m.Attachments.Count = 0 Then
        Log "Step1: Keyword found but no attachments."
        
        Dim r As VbMsgBoxResult
        r = MsgBox("本文に「添付」という文字がありますが、添付ファイルがありません。" & vbCrLf & _
                   "このまま送信しますか？", vbYesNo + vbExclamation, "Step1: 添付確認")
        
        Log "Step1 User Selection: " & MsgBoxResultToJa(r)
        
        If r = vbNo Then
            m.Save
            CheckAttachmentMention = False ' 中止
        Else
            allowNoAttachment = True ' 添付なしを許可
        End If
    Else
        Log "Step1: Skipped."
    End If
    
    Log "Step1 Time(ms): " & FormatTime(t0)
End Function

' ==============================================================================
' [Logic 2] Zip/7z パスワード確認ロジック
' ==============================================================================
Private Function CheckZipPassword(ByRef m As Outlook.MailItem) As Boolean
    CheckZipPassword = True
    
    Dim t0 As Double: t0 = Timer
    Dim tempFolder As String: tempFolder = Environ$("TEMP") & "\"
    Dim zipFiles As Collection: Set zipFiles = New Collection
    Dim at As Outlook.Attachment
    Dim tempPath As String
    
    ' 同名ファイル衝突回避用の連番カウンタ
    Dim i As Long
    i = 1
    
    ' 対象ファイルの抽出
    For Each at In m.Attachments
        If IsZipOr7z(at.FileName) Then
            ' 時刻 + 連番(i) を付与して一意性を確保
            tempPath = tempFolder & "chk_" & g_RunId & "_" & i & "_" & at.FileName
            
            at.SaveAsFile tempPath
            zipFiles.Add tempPath
            Log "Step2 Target: " & tempPath
            
            i = i + 1
        End If
    Next at
    
    If zipFiles.Count = 0 Then Exit Function
    
    ' 7-Zipによる判定
    Dim f As Variant
    Dim isEncrypted As Boolean
    Dim r As VbMsgBoxResult
    
    For Each f In zipFiles
        isEncrypted = CheckArchiveEncryption(CStr(f))
        
        If Not isEncrypted Then
            Log "Step2 Warning: No Password -> " & CStr(f)
            r = MsgBox("パスワードなしのZIP/7zが含まれています。" & vbCrLf & _
                       "ファイル: " & Dir(CStr(f)) & vbCrLf & vbCrLf & _
                       "送信を続けますか？", vbYesNo + vbExclamation, "Step2: セキュリティ確認")
            
            Log "Step2 User Selection: " & MsgBoxResultToJa(r)
            
            If r = vbNo Then
                m.Save
                CheckZipPassword = False
                GoTo CLEANUP
            End If
        End If
    Next f

CLEANUP:
    ' 一時ファイル削除
    On Error Resume Next
    For Each f In zipFiles
        Kill CStr(f)
    Next f
    On Error GoTo 0
    Log "Step2 Time(ms): " & FormatTime(t0)
End Function

Private Function IsZipOr7z(ByVal fileName As String) As Boolean
    Dim lower As String: lower = LCase$(fileName)
    IsZipOr7z = (Right$(lower, 4) = ".zip") Or (Right$(lower, 3) = ".7z")
End Function

Private Function CheckArchiveEncryption(ByVal path As String) As Boolean
    On Error GoTo EH
    
    ' modLoggerから共通の7-Zipパスを取得
    Dim sevenZipPath As String
    sevenZipPath = modLogger.GetSevenZipPath()
    
    ' クォート処理
    If Left(sevenZipPath, 1) <> """" Then sevenZipPath = """" & sevenZipPath & """"
    
    Dim tempFile As String, logFile As String
    ' g_RunIdを利用してファイル名競合を防止
    tempFile = Environ$("TEMP") & "\7zOut_" & g_RunId & ".txt"
    logFile = Environ$("TEMP") & "\7zErr_" & g_RunId & ".txt"
    
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    ' cmd.exe /s /c で引用符構文を正規化し、-p"" と < nul を付与して対話プロンプト待ちを完全に防止
    Dim cmd As String
    cmd = "cmd.exe /s /c """ & sevenZipPath & " l -slt """ & path & """ -p"""" < nul > """ & tempFile & """ 2> """ & logFile & """"
    
    Dim result As Long
    result = shell.Run(cmd, 0, True)
    Set shell = Nothing
    
    If result <> 0 Then
        ' エラー内容(標準エラー出力)を読み取ってログに出す
        Dim errText As String
        errText = ReadAllText(logFile)
        
        Log "7z Command Failed. Code=" & result
        Log "7z Error Details: " & Replace(Replace(errText, vbCrLf, " "), vbCr, " ")
        
        ' ヘッダー暗号化(.7zの-mheなど)によりパスワードなしで一覧が開けない場合、暗号化されていると正しく判定
        If InStr(1, errText, "encrypted archive", vbTextCompare) > 0 Or _
           InStr(1, errText, "Wrong password", vbTextCompare) > 0 Or _
           InStr(1, errText, "Headers Error", vbTextCompare) > 0 Or _
           InStr(1, errText, "暗号化", vbTextCompare) > 0 Then
            Log "Detected Header-Encrypted Archive (7z with -mhe)."
            CheckArchiveEncryption = True
        Else
            CheckArchiveEncryption = False ' その他のエラー時は安全側に倒す
        End If
        GoTo CLEANUP_FILES
    End If
    
    ' 結果解析
    Dim output As String
    output = ReadAllText(tempFile)
    
    ' 日本語・英語環境対応 (暗号化 = + / Encrypted = +)
    If InStr(1, output, "Encrypted = +", vbTextCompare) > 0 Or _
       InStr(1, output, "暗号化 = +", vbTextCompare) > 0 Then
        CheckArchiveEncryption = True
    Else
        CheckArchiveEncryption = False
    End If
    
CLEANUP_FILES:
    On Error Resume Next
    If Dir(tempFile) <> "" Then Kill tempFile
    If Dir(logFile) <> "" Then Kill logFile
    Exit Function
EH:
    Log "7z Error: " & Err.Description
    Resume CLEANUP_FILES
End Function

' ==============================================================================
' [Logic 3] 送信時刻制御ロジック
' ==============================================================================
Private Function CheckSendTime(ByRef m As Outlook.MailItem) As Boolean
    CheckSendTime = True
    Dim t0 As Double: t0 = Timer
    
    Dim nowTime As Date: nowTime = Now
    Dim isBizDay As Boolean: isBizDay = IsBusinessDay(nowTime)
    
    Dim t As Date: t = TimeValue(nowTime)
    
    ' 設定ファイルから業務時間を取得（デフォルト 08:00 ～ 18:00）
    Dim startTimeStr As String: startTimeStr = GetConfigValue("WorkStartTime", "08:00")
    Dim endTimeStr As String: endTimeStr = GetConfigValue("WorkEndTime", "18:00")
    
    ' 秒数が省略されている場合は付与
    If Len(startTimeStr) = 5 Then startTimeStr = startTimeStr & ":00"
    If Len(endTimeStr) = 5 Then endTimeStr = endTimeStr & ":00"
    
    Dim startVal As Date, endVal As Date
    On Error Resume Next
    startVal = TimeValue(startTimeStr)
    If Err.Number <> 0 Then startVal = TimeValue("08:00:00")
    Err.Clear
    endVal = TimeValue(endTimeStr)
    If Err.Number <> 0 Then endVal = TimeValue("18:00:00")
    On Error GoTo 0
    
    ' 時間外判定
    Dim isLate As Boolean: isLate = (t >= endVal)
    Dim isEarly As Boolean: isEarly = (t < startVal)
    
    Dim deferAt As Date
    Dim needConfirmation As Boolean
    
    If isLate Or isEarly Then
        deferAt = CalcDeferTime(nowTime, startVal, endVal)
        needConfirmation = True
        Log "Step3 Status: Night/Early. Candidate=" & deferAt
    ElseIf Not isBizDay Then
        deferAt = NextBusinessDayAtTime(nowTime, startVal)
        needConfirmation = True
        Log "Step3 Status: Holiday. Candidate=" & deferAt
    Else
        Log "Step3 Status: Normal business hours."
        Exit Function
    End If
    
    If needConfirmation Then
        Dim r As VbMsgBoxResult
        r = MsgBox("送信タイミングの確認" & vbCrLf & vbCrLf & _
                   "Yes : 翌営業日に予約送信 (" & Format(deferAt, "mm/dd hh:nn") & ")" & vbCrLf & _
                   "No  : 即時送信" & vbCrLf & _
                   "Cancel : 送信中止", _
                   vbYesNoCancel + vbQuestion, "Step3: 送信確認")
        
        Log "Step3 User Selection: " & MsgBoxResultToJa(r)
        
        Select Case r
            Case vbYes
                m.DeferredDeliveryTime = deferAt
                m.Save
                ' 予約して送信プロセス継続（Outboxへ）
            Case vbNo
                ' 即時送信
            Case vbCancel
                m.Save
                CheckSendTime = False
        End Select
    End If
    
    Log "Step3 Time(ms): " & FormatTime(t0)
End Function

' --- 日付計算ヘルパー ---
Private Function IsBusinessDay(ByVal d As Date) As Boolean
    Dim w As Long: w = Weekday(d, vbMonday)
    If w >= 6 Then
        IsBusinessDay = False
    Else
        IsBusinessDay = Not IsHoliday(d)
    End If
End Function

Private Function IsHoliday(ByVal d As Date) As Boolean
    Dim md As String: md = Format$(d, "mm-dd")
    Dim listStr As String
    listStr = GetConfigValue("HolidayList", DEFAULT_HOLIDAYS)
    
    Dim holidays() As String
    holidays = Split(listStr, ",")
    
    Dim i As Long
    For i = LBound(holidays) To UBound(holidays)
        If Trim$(holidays(i)) = md Then
            IsHoliday = True
            Exit Function
        End If
    Next i
    IsHoliday = False
End Function

' 指定した日付から、翌営業日の指定時刻(timeVal)を算出する
Private Function NextBusinessDayAtTime(ByVal baseDate As Date, ByVal timeVal As Date) As Date
    Dim d As Date: d = DateValue(baseDate) + 1
    Do While Not IsBusinessDay(d)
        d = d + 1
    Loop
    NextBusinessDayAtTime = d + timeVal
End Function

' 現在時刻と業務時間設定に基づき、予約送信時刻を算出する
Private Function CalcDeferTime(ByVal baseTime As Date, ByVal startVal As Date, ByVal endVal As Date) As Date
    Dim today As Date: today = DateValue(baseTime)
    Dim t As Date: t = TimeValue(baseTime)
    
    If t >= endVal Then
        CalcDeferTime = NextBusinessDayAtTime(today, startVal)
    ElseIf t < startVal Then
        If IsBusinessDay(today) Then
            CalcDeferTime = today + startVal
        Else
            CalcDeferTime = NextBusinessDayAtTime(today, startVal)
        End If
    Else
        CalcDeferTime = NextBusinessDayAtTime(today, startVal)
    End If
End Function

' ==============================================================================
' [Config] 設定読み込み (modConfig 連携)
' ==============================================================================

Private Function GetConfigValue(ByVal key As String, Optional ByVal defaultVal As String = "") As String
    Dim val As String
    val = modConfig.GetConfigValue("SendController", key, "")
    If Len(val) > 0 Then
        GetConfigValue = val
    Else
        If Len(defaultVal) > 0 Then
            Log "Config Key Not Found: [" & key & "] -> Using Default: " & defaultVal
        End If
        GetConfigValue = defaultVal
    End If
End Function

' ==============================================================================
' [Common] ヘルパー関数
' ==============================================================================

Private Function ReadAllText(ByVal filePath As String) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        With fso.OpenTextFile(filePath, 1)
            ReadAllText = .ReadAll
            .Close
        End With
    End If
End Function

Private Sub SaveToDraftsSafe(ByVal Item As Object)
    On Error Resume Next
    If TypeOf Item Is Outlook.MailItem Then
        Dim m As Outlook.MailItem: Set m = Item
        Dim ns As Outlook.NameSpace: Set ns = Application.GetNamespace("MAPI")
        Dim drafts As Outlook.MAPIFolder: Set drafts = ns.GetDefaultFolder(olFolderDrafts)
        Dim cp As Outlook.MailItem: Set cp = m.Copy
        Set cp = cp.Move(drafts)
        cp.Save
    End If
End Sub

Private Function MsgBoxResultToJa(ByVal r As VbMsgBoxResult) As String
    Select Case r
        Case vbYes: MsgBoxResultToJa = "Yes"
        Case vbNo: MsgBoxResultToJa = "No"
        Case vbCancel: MsgBoxResultToJa = "Cancel"
        Case Else: MsgBoxResultToJa = "Unknown"
    End Select
End Function

