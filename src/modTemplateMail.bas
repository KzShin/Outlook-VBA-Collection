Attribute VB_Name = "modTemplateMail"
Option Explicit

' ==============================================================================
' Module: modTemplateMail
' Description: 指定フォルダのテンプレート(.oft/.msg)から日付変数を置換して新規テキストメールを作成する
' Dependencies: Scripting.FileSystemObject, ADODB.Stream, modLogger, frmSelectTemplate
' Configuration: %APPDATA%\OutlookVBA\config.ini ([TemplateMail] Section)
' ==============================================================================

Private g_RunId As String

Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

Private Sub Log(ByVal msg As String)
    modLogger.Log "TemplateMail", msg
End Sub

' ==============================================================================
' [Public] 公開インターフェース
' ==============================================================================

Public Sub CreateMailFromTemplate()
    On Error GoTo EH
    
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-TMPL"
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START テンプレート作成処理 ==="
    
    ' 1. テンプレートフォルダのパスを取得・確認
    Dim tmplFolder As String
    tmplFolder = GetTemplateFolderPath()
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' フォルダが存在しなければ作成して終了
    If Not fso.FolderExists(tmplFolder) Then
        CreateFolderRecursive fso, tmplFolder
        MsgBox "テンプレート用フォルダを作成しました。" & vbCrLf & _
               "以下のフォルダに .oft または .msg ファイルを配置してから再度実行してください。" & vbCrLf & vbCrLf & _
               tmplFolder, vbInformation, "TemplateMail"
        Log "フォルダ新規作成: " & tmplFolder
        GoTo FIN
    End If
    
    ' 2. フォルダ内のテンプレートファイルを収集
    Dim file As Object
    Dim hasFiles As Boolean
    Dim frm As frmSelectTemplate
    Set frm = New frmSelectTemplate
    
    hasFiles = False
    For Each file In fso.GetFolder(tmplFolder).Files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(file.Name))
        If ext = "oft" Or ext = "msg" Then
            ' リストボックスに「ファイル名」と「フルパス」を登録
            frm.lstTemplates.AddItem file.Name
            frm.lstTemplates.List(frm.lstTemplates.ListCount - 1, 1) = file.Path
            hasFiles = True
        End If
    Next file
    
    If Not hasFiles Then
        MsgBox "指定されたフォルダにテンプレート(.oft または .msg)が見つかりません。" & vbCrLf & _
               tmplFolder, vbExclamation, "TemplateMail"
        Log "テンプレートファイルなし"
        Unload frm
        GoTo FIN
    End If
    
    ' 3. ユーザーに選択させる
    Log "選択ダイアログ表示"
    frm.Show
    
    If frm.IsCancelled Then
        Log "ユーザーキャンセル"
        Unload frm
        GoTo FIN
    End If
    
    Dim targetPath As String
    targetPath = frm.SelectedFilePath
    Unload frm
    
    Log "選択されたテンプレート: " & targetPath
    
    ' 4. メールアイテムの生成と日付置換
    Dim newItem As Object
    Set newItem = Application.CreateItemFromTemplate(targetPath)
    
    If TypeName(newItem) = "MailItem" Then
        Dim mailItem As Outlook.MailItem
        Set mailItem = newItem
        
        ' プレースホルダーの置換処理とテキスト形式への強制変換
        ReplaceDatePlaceholders mailItem
        
        Log "置換・テキスト形式変換完了、メールを表示"
        mailItem.Display
    Else
        MsgBox "選択されたファイルはメール形式のテンプレートではありません。", vbCritical
        Log "非対応アイテム: " & TypeName(newItem)
        newItem.Close olDiscard
    End If
    
FIN:
    Log "=== END テンプレート作成処理 ==="
    modLogger.SetRunId "NoID"
    Exit Sub
EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    modLogger.SetRunId "NoID"
End Sub

' ==============================================================================
' [Logic] 日付置換処理とテキスト形式変換
' ==============================================================================

Private Sub ReplaceDatePlaceholders(ByRef mail As Outlook.MailItem)
    Dim dtNow As Date: dtNow = Now
    
    ' 置換用文字列の生成（mm, dd は0埋めなし）
    Dim sYyyy As String: sYyyy = CStr(Year(dtNow))
    Dim sYy As String: sYy = Right$(sYyyy, 2)
    Dim sMm As String: sMm = CStr(Month(dtNow))
    Dim sDd As String: sDd = CStr(Day(dtNow))
    
    ' 和暦（令和〇）の計算
    Dim reiwaYear As Integer
    Dim sGgge As String
    reiwaYear = Year(dtNow) - 2018
    If reiwaYear = 1 Then
        sGgge = "令和元"
    Else
        sGgge = "令和" & CStr(reiwaYear)
    End If
    
    ' 曜日（月、火...）の取得 ※環境依存を避けるため配列で直接指定
    Dim sAaa As String
    sAaa = Choose(Weekday(dtNow), "日", "月", "火", "水", "木", "金", "土")
    
    ' 1. 件名の置換
    Dim subj As String
    subj = mail.Subject
    If Len(subj) > 0 Then
        subj = Replace(subj, "{yyyy}", sYyyy)
        subj = Replace(subj, "{yy}", sYy)
        subj = Replace(subj, "{mm}", sMm)
        subj = Replace(subj, "{dd}", sDd)
        subj = Replace(subj, "{ggge}", sGgge)
        subj = Replace(subj, "{aaa}", sAaa)
        mail.Subject = subj
    End If
    
    ' 2. テキスト形式に強制変換 (olFormatPlain = 1)
    mail.BodyFormat = olFormatPlain
    
    ' 3. 本文の置換 (テキスト形式化された Body に対して実行)
    Dim plain As String
    plain = mail.Body
    If Len(plain) > 0 Then
        plain = Replace(plain, "{yyyy}", sYyyy)
        plain = Replace(plain, "{yy}", sYy)
        plain = Replace(plain, "{mm}", sMm)
        plain = Replace(plain, "{dd}", sDd)
        plain = Replace(plain, "{ggge}", sGgge)
        plain = Replace(plain, "{aaa}", sAaa)
        mail.Body = plain
    End If
End Sub

' ==============================================================================
' [Config] 設定管理・フォルダ操作
' ==============================================================================

Private Function GetTemplateFolderPath() As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim configPath As String
    configPath = Environ$("APPDATA") & "\OutlookVBA\config.ini"
    
    ' デフォルトパス
    Dim defaultPath As String
    defaultPath = Environ$("APPDATA") & "\OutlookVBA\Templates"
    
    If Not fso.FileExists(configPath) Then
        GetTemplateFolderPath = defaultPath
        Exit Function
    End If
    
    ' UTF-8 読み込み
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8": stm.Open
    stm.LoadFromFile configPath
    Dim allText As String: allText = stm.ReadText(-1)
    stm.Close
    
    Dim lines() As String: lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    Dim i As Long, lineText As String, currentSection As String, parts() As String
    
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
            If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
            ElseIf currentSection = "templatemail" Then
                parts = Split(lineText, "=", 2)
                If UBound(parts) = 1 Then
                    If Trim$(LCase$(parts(0))) = "templatefolder" Then
                        Dim val As String
                        val = Trim$(parts(1))
                        ' %APPDATA% の環境変数を展開
                        val = Replace(val, "%APPDATA%", Environ$("APPDATA"), 1, -1, vbTextCompare)
                        GetTemplateFolderPath = val
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    
    GetTemplateFolderPath = defaultPath
End Function

Private Sub CreateFolderRecursive(ByVal fso As Object, ByVal path As String)
    Dim p As String: p = fso.GetParentFolderName(path)
    If Not fso.FolderExists(p) Then CreateFolderRecursive fso, p
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub
