Attribute VB_Name = "modAutoFlag"
Option Explicit

' ==============================================================================
' Module: modAutoFlag
' Description: 受信メールを解析し、条件（件名、宛先、本文）に応じて自動的にフラグを設定する
' Dependencies: Scripting.FileSystemObject, VBScript.RegExp, ADODB.Stream, modLogger
' Configuration: %APPDATA%\OutlookVBA\config.ini ([AutoFlag] Section)
' ==============================================================================

' --- ThisOutlookSessionでの呼び出し例 ---
' Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'     modAutoFlag.ProcessNewMail EntryIDCollection
' End Sub
' Private Sub Application_Startup()
'     modAutoFlag.ProcessStartupUnread
' End Sub

' --- グローバル変数 ---
Private g_RunId As String

' ==============================================================================
' [Public] 初期化・ログ設定
' ==============================================================================

' 実行IDを設定
Public Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

' 共通ロガーへの委譲
Private Sub Log(ByVal msg As String)
    modLogger.Log "AutoFlag", msg
End Sub

' ==============================================================================
' [Main] メイン処理
' ==============================================================================

' 1. 受信時処理 (NewMailExから呼び出し)
Public Sub ProcessNewMail(ByVal EntryIDCollection As String)
    On Error GoTo EH
    
    ' 実行ID生成
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-NEW"
    
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START メール自動フラグ処理 (受信) ==="

    ' 設定とRegExpの準備
    Dim myAddress As String, regexPattern As String, excludeSubjects As Variant
    Dim RegEx As Object
    
    If Not PrepareConfigAndRegEx(myAddress, regexPattern, excludeSubjects, RegEx) Then
        Exit Sub
    End If

    ' メール取得ループ
    Dim olNs As Outlook.NameSpace
    Dim olItem As Object
    Dim EntryIDs() As String
    Dim i As Integer
    
    Set olNs = Application.GetNamespace("MAPI")
    EntryIDs = Split(EntryIDCollection, ",")
    
    For i = 0 To UBound(EntryIDs)
        On Error Resume Next
        Set olItem = olNs.GetItemFromID(EntryIDs(i))
        On Error GoTo EH
        
        ' 共通処理へ渡す
        Call ProcessSingleMail(olItem, myAddress, excludeSubjects, RegEx)
        Set olItem = Nothing
    Next i

    Log "=== END メール自動フラグ処理 (受信) ==="
    Exit Sub
EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
End Sub

' 2. 起動時処理 (Startupから呼び出し)
Public Sub ProcessStartupUnread()
    On Error GoTo EH
    
    ' 実行ID生成
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-BOOT"
    
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START メール自動フラグ処理 (起動時未読チェック) ==="

    ' 設定とRegExpの準備
    Dim myAddress As String, regexPattern As String, excludeSubjects As Variant
    Dim RegEx As Object
    
    If Not PrepareConfigAndRegEx(myAddress, regexPattern, excludeSubjects, RegEx) Then
        Exit Sub
    End If

    ' 受信トレイの未読メールを取得
    Dim olNs As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim items As Outlook.Items
    Dim mailItem As Object
    
    Set olNs = Application.GetNamespace("MAPI")
    Set inbox = olNs.GetDefaultFolder(olFolderInbox)
    
    ' 未読のみフィルタリング
    Set items = inbox.items.Restrict("[UnRead] = True")
    Log "受信トレイ未読件数: " & items.Count & "件"
    
    ' 未読アイテムループ
    For Each mailItem In items
        On Error Resume Next
        ' 共通処理へ渡す
        Call ProcessSingleMail(mailItem, myAddress, excludeSubjects, RegEx)
        On Error GoTo EH
    Next

    Log "=== END メール自動フラグ処理 (起動時未読チェック) ==="
    Exit Sub
EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
End Sub


' ==============================================================================
' [Private] 共通ロジック
' ==============================================================================

' 設定読み込みと正規表現生成をまとめた関数
Private Function PrepareConfigAndRegEx(ByRef outAddress As String, _
                                       ByRef outPattern As String, _
                                       ByRef outExcludes As Variant, _
                                       ByRef outRegEx As Object) As Boolean
    ' 設定読み込み
    If Not LoadConfig(outAddress, outPattern, outExcludes) Then
        Log "設定読み込み失敗または必須項目不足により中止"
        PrepareConfigAndRegEx = False
        Exit Function
    End If
    
    ' 正規表現生成
    Set outRegEx = CreateObject("VBScript.RegExp")
    With outRegEx
        .Pattern = outPattern
        .IgnoreCase = True
        .Global = False
    End With
    
    Log "設定準備完了: Pattern=" & outPattern & ", Excludes=" & (UBound(outExcludes) + 1) & "件"
    PrepareConfigAndRegEx = True
End Function

' 1通のメールに対する判定とフラグ付与 (共通処理)
Private Sub ProcessSingleMail(ByVal objItem As Object, _
                              ByVal myAddress As String, _
                              ByVal excludeSubjects As Variant, _
                              ByVal RegEx As Object)
    If objItem Is Nothing Then Exit Sub
    
    ' MailItem以外は除外
    If TypeName(objItem) <> "MailItem" Then Exit Sub
    
    Dim olMail As Outlook.MailItem
    Set olMail = objItem
    
    ' A. 除外判定
    If IsExcludedSubject(olMail.Subject, excludeSubjects) Then
        Log "スキップ(除外KW): " & olMail.Subject
        Exit Sub
    End If
    
    ' B. 条件一致判定
    If CheckMatchConditions(olMail, myAddress, RegEx) Then
        ' C. フラグ設定
        If olMail.FlagStatus <> olFlagMarked Then
            With olMail
                .FlagRequest = "要確認"
                .FlagStatus = olFlagMarked
                .Save
            End With
            Log ">>> フラグ設定: " & olMail.Subject
        Else
            Log "フラグ設定済(スキップ): " & olMail.Subject
        End If
    End If
End Sub


' ==============================================================================
' [Config] 設定管理
' ==============================================================================

Private Function LoadConfig(ByRef outAddress As String, _
                            ByRef outPattern As String, _
                            ByRef outExcludes As Variant) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 変更: 共通設定ファイルのパス
    Dim configPath As String
    configPath = Environ("APPDATA") & "\OutlookVBA\config.ini"
    
    If Not fso.FileExists(configPath) Then
        Log "設定ファイルなし: " & configPath
        LoadConfig = False
        Exit Function
    End If
    
    outAddress = ""
    outPattern = ""
    outExcludes = Array()
    
    On Error GoTo EH
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .LoadFromFile configPath
    End With
    
    Dim allText As String
    allText = stm.ReadText(-1)
    stm.Close
    
    Dim lines() As String
    lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    
    Dim i As Long, lineText As String, parts() As String
    Dim currentSection As String
    
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        ' コメントと空行スキップ
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
            
            ' セクション判定 [SectionName]
            If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
            
            ' [AutoFlag] セクションのみ処理
            ElseIf currentSection = "autoflag" Then
                parts = Split(lineText, "=", 2)
                If UBound(parts) = 1 Then
                    Select Case Trim(parts(0))
                        Case "MyAddress"
                            outAddress = Trim(parts(1))
                        Case "Pattern"
                            outPattern = Trim(parts(1))
                        Case "ExcludeSubjects"
                            outExcludes = Split(parts(1), ",")
                    End Select
                End If
            End If
            
        End If
    Next i
    
    If Len(outAddress) = 0 Then
        Log "Config Error: MyAddress未定義"
        LoadConfig = False
    Else
        LoadConfig = True
    End If
    Exit Function
EH:
    Log "LoadConfig Error: " & Err.Description
    LoadConfig = False
End Function


' ==============================================================================
' [Logic] 判定ロジック
' ==============================================================================

Private Function IsExcludedSubject(ByVal subject As String, ByVal excludeList As Variant) As Boolean
    Dim keyword As Variant
    IsExcludedSubject = False
    If Not IsArray(excludeList) Then Exit Function
    
    For Each keyword In excludeList
        If Len(keyword) > 0 Then
            If InStr(1, subject, Trim(keyword), vbTextCompare) > 0 Then
                IsExcludedSubject = True
                Exit Function
            End If
        End If
    Next keyword
End Function

Private Function CheckMatchConditions(ByVal mail As Outlook.MailItem, _
                                      ByVal myAddress As String, _
                                      ByVal regExObj As Object) As Boolean
    Dim toMatches As Boolean, ccMatches As Boolean, bodyMatches As Boolean
    
    toMatches = (InStr(1, mail.To, myAddress, vbTextCompare) > 0)
    ccMatches = (InStr(1, mail.CC, myAddress, vbTextCompare) > 0)
    bodyMatches = regExObj.Test(mail.Body)
    
    CheckMatchConditions = (toMatches Or ccMatches Or bodyMatches)
End Function
