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
Private g_ConfigCache As Object     ' 設定キャッシュ (Dictionary)
Private g_IsConfigLoaded As Boolean ' 設定読み込み済みフラグ

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
    
    ' 3. 設定読み込み (初回のみ)
    If Not g_IsConfigLoaded Then LoadConfig
    
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
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "送信処理中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    Cancel = True
    ' エラー時は念のため下書きへ退避
    SaveToDraftsSafe Item
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
    FormatTime = CStr(CLng((Timer - startTime) * 1000))
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
    
    ' 変更: modLoggerから共通の7-Zipパスを取得
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
    
    ' cmd.exe /c の引用符剥がれ対策として全体を囲む
    Dim cmd As String
    cmd = "cmd.exe /c """ & sevenZipPath & " l -slt """ & path & """ > """ & tempFile & """ 2> """ & logFile & """"" "
    
    Dim result As Long
    result = shell.Run(cmd, 0, True)
    
    If result <> 0 Then
        ' エラー内容(標準エラー出力)を読み取ってログに出す
        Dim errText As String
        errText = ReadAllText(logFile)
        
        Log "7z Command Failed. Code=" & result
        Log "7z Error Details: " & Replace(Replace(errText, vbCrLf, " "), vbCr, " ")
        
        CheckArchiveEncryption = False ' エラー時は安全側に倒す
        GoTo CLEANUP_FILES
    End If
    
    ' 結果解析
    Dim output As String
    output = ReadAllText(tempFile)
    
    ' 日本語環境対応 (暗号化 = +)
    If InStr(output, "Encrypted = +") > 0 Or InStr(output, "暗号化 = +") > 0 Then
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
    Dim isLate As Boolean: isLate = (t >= TimeValue("18:00:00"))
    Dim isEarly As Boolean: isEarly = (t < TimeValue("08:00:00"))
    
    Dim deferAt As Date
    Dim needConfirmation As Boolean
    
    If isLate Or isEarly Then
        deferAt = CalcDeferTime(nowTime)
        needConfirmation = True
        Log "Step3 Status: Night/Early. Candidate=" & deferAt
    ElseIf Not isBizDay Then
        deferAt = NextBusinessDayAt8FromDate(nowTime)
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

Private Function NextBusinessDayAt8FromDate(ByVal baseDate As Date) As Date
    Dim d As Date: d = DateValue(baseDate) + 1
    Do While Not IsBusinessDay(d)
        d = d + 1
    Loop
    NextBusinessDayAt8FromDate = d + TimeValue("08:00:00")
End Function

Private Function CalcDeferTime(ByVal baseTime As Date) As Date
    Dim today As Date: today = DateValue(baseTime)
    Dim t As Date: t = TimeValue(baseTime)
    
    If t >= TimeValue("18:00:00") Then
        CalcDeferTime = NextBusinessDayAt8FromDate(today)
    ElseIf t < TimeValue("08:00:00") Then
        If IsBusinessDay(today) Then
            CalcDeferTime = today + TimeValue("08:00:00")
        Else
            CalcDeferTime = NextBusinessDayAt8FromDate(today)
        End If
    Else
        CalcDeferTime = NextBusinessDayAt8FromDate(today)
    End If
End Function

' ==============================================================================
' [Config] 設定読み込み (統合版 config.ini 対応)
' ==============================================================================
Private Sub LoadConfig()
    On Error GoTo EH
    Set g_ConfigCache = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim configPath As String
    ' 統合設定ファイルのパス
    configPath = Environ("APPDATA") & "\OutlookVBA\config.ini"
    
    If Not fso.FileExists(configPath) Then
        Log "Config file not found: " & configPath
        g_IsConfigLoaded = True
        Exit Sub
    End If
    
    ' ADODB.StreamによるUTF-8読み込み
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2 ' adTypeText
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
            
            ' [SendController] セクションのみ読み込む
            ElseIf currentSection = "sendcontroller" Then
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
        ' キーが見つからずデフォルト値を使う場合にログ出力
        Log "Config Key Not Found: [" & key & "] -> Using Default: " & defaultVal
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
