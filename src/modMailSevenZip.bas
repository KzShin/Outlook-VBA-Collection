Attribute VB_Name = "modMailSevenZip"
Option Explicit

' ==============================================================================
' Module: modMailSevenZip
' Description: メール添付ファイルを保存し、7-Zipを使用して自動解凍を行うモジュール
' Dependencies: 7-Zip, modLogger
' Configuration: %APPDATA%\OutlookVBA\SevenZipPasswords.txt (for Password List)
'                %APPDATA%\OutlookVBA\config.ini ([General] Section)
' ==============================================================================

' --- ThisOutlookSessionでの呼び出し例 ---
' Public Sub SaveMailAttachmentsMacro()
'     Dim rid As String: rid = Format(Now, "yymmdd-hhnnss") & "-SAVE"
'     modLogger.SetRunId rid
'     modMailSevenZip.SaveAndExtractAttachments
' End Sub

' ==============================================================================
' [Private] ログ出力ヘルパー（modLoggerへの委譲）
' ==============================================================================

Private Sub Log(ByVal msg As String)
    modLogger.Log "MailSevenZip", msg
End Sub


' ==============================================================================
' [Main] メイン処理
' ==============================================================================

' 選択したメールを保存し、アーカイブであれば解凍を試みる
' 変更: メソッド名を英語(PascalCase)に統一
Public Sub SaveAndExtractAttachments()
    On Error GoTo EH
    Log "=== START メール保存・解凍処理 ==="
    
    ' 1. メール選択チェック
    Dim objItem As Object
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "メールを選択してください。", vbExclamation
        Log "選択なし：処理終了"
        Exit Sub
    End If
    
    Set objItem = Application.ActiveExplorer.Selection(1)
    If objItem.Class <> olMail Then
        MsgBox "選択されたアイテムはメールではありません。", vbExclamation
        Log "非メールアイテム：Class=" & objItem.Class & " : 処理終了"
        Exit Sub
    End If
    
    Dim mail As Outlook.MailItem
    Set mail = objItem
    Log "対象メール：Subject=""" & mail.Subject & """ / Received=" & mail.ReceivedTime
    
    ' 2. 保存先パスの構築
    Dim receivedDate As String
    receivedDate = Format(mail.ReceivedTime, "yymmdd_hhnnss")
    
    Dim safeSubject As String
    safeSubject = SafeName(mail.Subject)
    
    Dim rootPath As String
    rootPath = Environ$("USERPROFILE") & "\Downloads\" & receivedDate & "_" & safeSubject & "\"
    Log "保存先ルート：" & rootPath
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 既存フォルダチェック
    If fso.FolderExists(rootPath) Then
        MsgBox "保存先フォルダが既に存在します。既存フォルダを開きます。" & vbCrLf & rootPath, vbInformation
        Log "既存フォルダ検出：処理中止 " & rootPath
        Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
        Exit Sub
    End If
    
    ' フォルダ作成
    fso.CreateFolder rootPath
    Log "フォルダ作成済み：" & rootPath
    
    ' 3. 7-Zip 環境確認 (共通モジュールの設定を利用)
    Dim sevenZipPath As String
    sevenZipPath = modLogger.GetSevenZipPath() 
    Dim has7zip As Boolean
    has7zip = (Len(sevenZipPath) > 0)
    
    If Not has7zip Then
        MsgBox "7-Zip が見つかりません。展開処理はスキップします。", vbExclamation
        Log "7-Zip未検出：展開スキップ"
    Else
        Log "7-Zip検出：" & sevenZipPath
    End If
    
    ' 4. 添付ファイルの保存と解凍処理
    Dim att As Outlook.Attachment
    For Each att In mail.Attachments
        Dim savePath As String
        savePath = fso.BuildPath(rootPath, att.FileName)
        savePath = MakeUniqueFilePath(fso, savePath)
        
        ' ファイル保存
        Log "添付保存開始：""" & att.FileName & """ -> " & savePath
        att.SaveAsFile savePath
        Log "添付保存完了：" & savePath
        
        ' アーカイブ処理（.zip / .7z）
        If has7zip And IsArchiveTarget(att.FileName) Then
            Dim baseName As String
            baseName = SafeName(fso.GetBaseName(att.FileName))
            
            Dim outDir As String
            outDir = fso.BuildPath(rootPath, baseName)
            outDir = MakeUniqueFolderPath(fso, outDir)
            fso.CreateFolder outDir
            Log "展開先フォルダ作成：" & outDir
            
            ' 解凍試行（タイムアウト 120秒）
            Dim ok As Boolean
            Dim cancelled As Boolean
            Log "事前テスト開始：" & savePath
            
            ok = TestThenExtractArchive(savePath, outDir, sevenZipPath, mail.ReceivedTime, 120, cancelled)
            
            If ok Then
                Log "展開成功：" & outDir
            Else
                Log "展開失敗：全候補不一致またはタイムアウト。後処理を実行"
                DeleteFolderIfEmpty fso, outDir
                
                If cancelled Then
                    Log "ユーザーキャンセルにより終了"
                Else
                    MsgBox "パスワード候補では解凍できませんでした: " & att.FileName & vbCrLf & _
                           "アーカイブは保存したままにしています。", vbInformation
                End If
            End If
        Else
            Log "非対象拡張子または7-Zipなし：保存のみ。 File=" & att.FileName
        End If
    Next att
    
    ' 5. メタデータ保存（メール本文など）
    Dim infoPath As String
    infoPath = fso.BuildPath(rootPath, "メール本文.txt")
    Log "本文書き出し開始：" & infoPath
    WriteTextUtf8 infoPath, BuildMailInfoText(mail)
    Log "本文書き出し完了：" & infoPath
    
    ' 6. 完了後のフォルダ表示
    Log "エクスプローラ起動：" & rootPath
    Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
    
    Log "=== END メール保存・解凍処理 ==="
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


' ==============================================================================
' [Core Logic] 解凍フロー制御
' ==============================================================================

' アーカイブ判定 (.zip / .7z)
Private Function IsArchiveTarget(ByVal fileName As String) As Boolean
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos = 0 Then
        IsArchiveTarget = False
        Exit Function
    End If
    
    Dim ext As String
    ext = LCase$(Mid$(fileName, dotPos + 1))
    
    Dim isTarget As Boolean
    isTarget = (ext = "zip" Or ext = "7z")
    
    Log "拡張子判定：" & ext & " -> " & IIf(isTarget, "対象", "非対象")
    IsArchiveTarget = isTarget
End Function

' テスト実行 → 解凍実行 の統合ロジック
Private Function TestThenExtractArchive(ByVal zipPath As String, ByVal outDir As String, _
                                        ByVal sevenZipPath As String, ByVal receivedTime As Date, _
                                        ByVal timeoutSeconds As Long, ByRef userCancelled As Boolean) As Boolean
    userCancelled = False
    Dim rc As Long

    ' 1) パスワード無しでのテスト
    Log "テスト（パス無し）実行"
    rc = SevenZipTest(zipPath, sevenZipPath, "", timeoutSeconds)
    
    If rc <= 1 Then
        Log "テスト成功（パス無し）。抽出実行へ"
        rc = SevenZipExtract(zipPath, outDir, sevenZipPath, "", timeoutSeconds)
        Log "抽出（パス無し）終了コード：" & rc
        
        PauseSeconds 0.5 ' ファイルシステム同期待ち
        If rc = 0 Or HasNonZeroFileDeep(outDir) Then
            TestThenExtractArchive = True
            Exit Function
        End If
        Log "抽出後チェック：ファイル未生成のため続行"
    End If

    ' 2) 登録済みパスワード候補でのテスト
    Dim cands As Collection
    Set cands = GetPasswordCandidates(receivedTime)
    Log "候補数：" & cands.Count

    Dim pw As Variant
    Dim idx As Long: idx = 0
    
    For Each pw In cands
        idx = idx + 1
        Log "テスト候補" & idx & " 試行"
        rc = SevenZipTest(zipPath, sevenZipPath, CStr(pw), timeoutSeconds)

        If rc <= 1 Then
            Log "テスト成功：候補一致 → 抽出へ"
            rc = SevenZipExtract(zipPath, outDir, sevenZipPath, CStr(pw), timeoutSeconds)
            Log "抽出終了コード（候補" & idx & "）： " & rc
            
            PauseSeconds 0.5 ' ファイルシステム同期待ち
            If rc = 0 Or HasNonZeroFileDeep(outDir) Then
                Log "抽出後チェック：成功"
                TestThenExtractArchive = True
                Exit Function
            Else
                Log "抽出後チェック：ファイル未生成のため次候補へ"
            End If
        End If
    Next pw

    ' 3) 全候補失敗 → ユーザー手入力へ移行
    Log "全候補失敗 → 手入力モード移行"
    If PromptAndTryPassword(zipPath, outDir, sevenZipPath, timeoutSeconds, userCancelled) Then
        TestThenExtractArchive = True
        Exit Function
    End If

    ' 解凍不可
    TestThenExtractArchive = False
End Function

' 手入力による解凍試行（成功またはキャンセルまでループ）
Private Function PromptAndTryPassword(ByVal zipPath As String, ByVal outDir As String, _
                                      ByVal sevenZipPath As String, ByVal timeoutSeconds As Long, _
                                      ByRef userCancelled As Boolean) As Boolean
    On Error GoTo EH
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pw As String
    Dim rc As Long

    userCancelled = False

    Do
        pw = InputBox( _
            Prompt:="登録済みパスワードでは解凍できませんでした。" & vbCrLf & _
                    "解凍用パスワードを入力してください（キャンセルで中止）。", _
            Title:="パスワード入力（7-Zip）" _
        )

        If Len(pw) = 0 Then
            Log "手入力：キャンセルまたは空入力のため中止"
            userCancelled = True
            PromptAndTryPassword = False
            Exit Function
        End If

        ' リトライのためフォルダをリセット
        On Error Resume Next
        If fso.FolderExists(outDir) Then fso.DeleteFolder outDir, True
        fso.CreateFolder outDir
        On Error GoTo EH

        ' テスト実行
        Log "手入力PWでテスト開始"
        rc = SevenZipTest(zipPath, sevenZipPath, pw, timeoutSeconds)

        If rc <= 1 Then
            Log "手入力PWで抽出開始"
            rc = SevenZipExtract(zipPath, outDir, sevenZipPath, pw, timeoutSeconds)
            
            PauseSeconds 0.5
            If rc = 0 Or HasNonZeroFileDeep(outDir) Then
                Log "手入力PW：成功"
                PromptAndTryPassword = True
                Exit Function
            Else
                Log "手入力PW：抽出後ファイル確認できず → 再入力"
            End If
        Else
            Log "手入力PW：テスト失敗 → 再入力"
        End If
    Loop

EH:
    Log "PromptAndTryPassword エラー: " & Err.Number & " " & Err.Description
    PromptAndTryPassword = False
End Function


' ==============================================================================
' [7-Zip Wrapper] コマンドライン実行
' ==============================================================================

' 7-Zip テスト (t command)
Private Function SevenZipTest(ByVal zipPath As String, ByVal sevenZipPath As String, _
                              ByVal password As String, ByVal timeoutSeconds As Long) As Long
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = """" & sevenZipPath & """ t -y " & _
          """" & zipPath & """" & _
          " -bso0 -bse0 -bsp0"

    ' パスワード指定（空でも -p"" を明示して対話モードを回避）
    cmd = cmd & " -p""" & password & """"

    Log "7zテスト起動：" & cmd
    Dim proc As Object
    Set proc = sh.Exec(cmd)

    SevenZipTest = WaitProcessWithTimeout(proc, timeoutSeconds)
    Log "7zテスト終了コード：" & SevenZipTest
End Function

' 7-Zip 抽出 (x command)
Private Function SevenZipExtract(ByVal zipPath As String, ByVal outDir As String, _
                                 ByVal sevenZipPath As String, ByVal password As String, _
                                 ByVal timeoutSeconds As Long) As Long
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim cmd As String
    cmd = """" & sevenZipPath & """ x -y " & _
          """" & zipPath & """ -o""" & outDir & """" & _
          " -bso0 -bse0 -bsp0"

    If Len(password) > 0 Then
        cmd = cmd & " -p""" & password & """"
    End If

    Log "7z抽出起動：" & cmd
    Dim proc As Object
    Set proc = sh.Exec(cmd)

    SevenZipExtract = WaitProcessWithTimeout(proc, timeoutSeconds)
    Log "7z抽出終了コード：" & SevenZipExtract
End Function


' ==============================================================================
' [Config] パスワード候補管理
' ==============================================================================

' 外部定義ファイルからパスワード候補を取得
' パス: %APPDATA%\OutlookVBA\SevenZipPasswords.txt
Private Function GetPasswordCandidates(ByVal receivedTime As Date) As Collection
    Dim col As New Collection
    Dim appData As String
    appData = Environ$("APPDATA")

    Dim folderPath As String
    ' OutlookVBAフォルダへ統合
    folderPath = appData & "\OutlookVBA"

    Dim listPath As String
    listPath = folderPath & "\SevenZipPasswords.txt"

    ' 設定フォルダがない場合は作成のみ行う
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
        Log "設定フォルダ作成: " & folderPath
    End If
    On Error GoTo 0

    ' ファイル読み込み
    Dim loaded As Boolean
    loaded = LoadPasswordsFromFile(listPath, receivedTime, col)

    If Not loaded Then
        Log "パスワードリスト読込なし（0件またはファイル未存在）: " & listPath
    End If

    Set GetPasswordCandidates = col
End Function

' ファイル読込と日付プレースホルダの展開
Private Function LoadPasswordsFromFile(ByVal filePath As String, _
                                       ByVal receivedTime As Date, _
                                       ByRef outCol As Collection) As Boolean
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        LoadPasswordsFromFile = False
        Exit Function
    End If

    ' 日付文字列の準備
    Dim yyyy As String, yy As String, mm As String, dd As String
    yyyy = Format(receivedTime, "yyyy")
    yy = Right$(yyyy, 2)
    mm = Format(receivedTime, "mm")
    dd = Format(receivedTime, "dd")

    ' UTF-8 で読み込み
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile filePath

    Dim allText As String
    allText = stm.ReadText(-1)
    stm.Close

    Dim lines() As String
    lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)

    Dim i As Long, raw As String, expanded As String, cnt As Long
    For i = LBound(lines) To UBound(lines)
        raw = Trim$(lines(i))
        ' 空行とコメント(#)をスキップ
        If Len(raw) > 0 And Left$(raw, 1) <> "#" Then
            expanded = raw
            expanded = Replace(expanded, "{yyyy}", yyyy)
            expanded = Replace(expanded, "{yy}", yy)
            expanded = Replace(expanded, "{mm}", mm)
            expanded = Replace(expanded, "{dd}", dd)

            outCol.Add expanded
            cnt = cnt + 1
            Log "候補追加：" & expanded
        End If
    Next i

    LoadPasswordsFromFile = (cnt > 0)
    Log "外部ファイル読込完了: " & cnt & "件"
    Exit Function

EH:
    Log "LoadPasswordsFromFile エラー: " & Err.Number & " " & Err.Description
    LoadPasswordsFromFile = False
End Function


' ==============================================================================
' [System] プロセス制御ユーティリティ
' ==============================================================================

' プロセス待機（タイムアウト付き）
Private Function WaitProcessWithTimeout(ByVal proc As Object, ByVal timeoutSeconds As Long) As Long
    Dim startTick As Single
    startTick = Timer

    Do
        ' 標準出力・エラーのバッファ消費
        Do While Not proc.StdOut.AtEndOfStream
            Dim s As String: s = proc.StdOut.Read(1024)
        Loop
        Do While Not proc.StdErr.AtEndOfStream
            Dim e As String: e = proc.StdErr.Read(1024)
        Loop

        If Not IsProcessRunning(proc) Then Exit Do

        ' タイムアウト監視
        If ElapsedSeconds(startTick) >= timeoutSeconds Then
            Log "タイムアウト発生：" & ElapsedSeconds(startTick) & "秒 (PID=" & proc.ProcessID & ")"
            On Error Resume Next
            TerminateProcessByPID proc.ProcessID
            On Error GoTo 0
            WaitProcessWithTimeout = 255
            Exit Function
        End If

        PauseSeconds 0.05
    Loop

    WaitProcessWithTimeout = proc.ExitCode
End Function

Private Function IsProcessRunning(ByVal proc As Object) As Boolean
    On Error Resume Next
    Dim code As Long: code = proc.ExitCode
    If Err.Number <> 0 Then
        Err.Clear: IsProcessRunning = True
    Else
        IsProcessRunning = False
    End If
    On Error GoTo 0
End Function

Private Sub TerminateProcessByPID(ByVal pid As Long)
    On Error Resume Next
    Dim svc As Object, obj As Object
    Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set obj = svc.Get("Win32_Process.Handle='" & pid & "'")
    If Not obj Is Nothing Then
        obj.Terminate
        Log "プロセス強制終了: PID=" & pid
    End If
    On Error GoTo 0
End Sub

' Timerの日跨ぎを考慮した経過秒数計算
Private Function ElapsedSeconds(ByVal startTick As Single) As Double
    Dim t As Double: t = Timer
    If t >= startTick Then
        ElapsedSeconds = t - startTick
    Else
        ElapsedSeconds = (86400# - startTick) + t
    End If
End Function

Private Sub PauseSeconds(ByVal seconds As Double)
    Dim st As Double: st = Timer
    Do While ElapsedSeconds(st) < seconds
        DoEvents
    Loop
End Sub


' ==============================================================================
' [Utils] ファイル・テキスト操作ユーティリティ
' ==============================================================================

' フォルダ内のファイル存在確認（再帰）
Private Function HasNonZeroFileDeep(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Dim fso As Object, fld As Object, f As Object, subf As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function

    Set fld = fso.GetFolder(folderPath)

    For Each f In fld.Files
        If f.Size > 0 Then
            HasNonZeroFileDeep = True
            Exit Function
        End If
    Next f

    For Each subf In fld.SubFolders
        If HasNonZeroFileDeep(subf.path) Then
            HasNonZeroFileDeep = True
            Exit Function
        End If
    Next subf
    On Error GoTo 0
End Function

Private Sub DeleteFolderIfEmpty(ByVal fso As Object, ByVal folderPath As String)
    On Error Resume Next
    If fso.FolderExists(folderPath) Then
        Dim fld As Object: Set fld = fso.GetFolder(folderPath)
        If fld.Files.Count = 0 And fld.SubFolders.Count = 0 Then
            fso.DeleteFolder folderPath, True
            Log "空フォルダ削除：" & folderPath
        End If
    End If
End Sub

' ファイル名に使えない文字を置換
Private Function SafeName(ByVal s As String) As String
    Dim r As String: r = s
    r = Replace(r, "<", "_")
    r = Replace(r, ">", "_")
    r = Replace(r, ":", "_")
    r = Replace(r, """", "_")
    r = Replace(r, "/", "_")
    r = Replace(r, "\", "_")
    r = Replace(r, "|", "_")
    r = Replace(r, "?", "_")
    r = Replace(r, "*", "_")
    r = Replace(r, vbCr, "_")
    r = Replace(r, vbLf, "_")
    r = Replace(r, vbTab, "_")
    r = Trim$(r)
    If Len(r) > 150 Then r = Left$(r, 150)
    SafeName = r
End Function

' 同名ファイルがある場合に連番を付与
Private Function MakeUniqueFilePath(ByVal fso As Object, ByVal path As String) As String
    If Not fso.FileExists(path) Then
        MakeUniqueFilePath = path
        Exit Function
    End If
    Dim folder As String, name As String, ext As String
    folder = fso.GetParentFolderName(path)
    name = fso.GetBaseName(path)
    ext = fso.GetExtensionName(path)
    
    Dim i As Long, cand As String: i = 2
    Do
        cand = fso.BuildPath(folder, name & " (" & i & ")." & ext)
        If Not fso.FileExists(cand) Then
            MakeUniqueFilePath = cand
            Exit Function
        End If
        i = i + 1
    Loop
End Function

' 同名フォルダがある場合に連番を付与
Private Function MakeUniqueFolderPath(ByVal fso As Object, ByVal path As String) As String
    If Not fso.FolderExists(path) Then
        MakeUniqueFolderPath = path
        Exit Function
    End If
    Dim folder As String, name As String
    folder = fso.GetParentFolderName(path)
    name = fso.GetFileName(path)
    
    Dim i As Long, cand As String: i = 2
    Do
        cand = fso.BuildPath(folder, name & " (" & i & ")")
        If Not fso.FolderExists(cand) Then
            MakeUniqueFolderPath = cand
            Exit Function
        End If
        i = i + 1
    Loop
End Function

' メタデータテキスト生成
Private Function BuildMailInfoText(ByVal mail As Outlook.MailItem) As String
    Dim sb As String
    sb = ""
    sb = sb & "受信日時: " & mail.ReceivedTime & vbCrLf
    sb = sb & "From: " & mail.SenderName & " <" & mail.SenderEmailAddress & ">" & vbCrLf
    sb = sb & "To: " & mail.To & vbCrLf
    sb = sb & "CC: " & mail.CC & vbCrLf
    sb = sb & "件名: " & mail.Subject & vbCrLf
    sb = sb & vbCrLf
    sb = sb & "メール本文:" & vbCrLf
    sb = sb & mail.Body & vbCrLf
    BuildMailInfoText = sb
End Function

' UTF-8 テキスト書き出し
Private Sub WriteTextUtf8(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2          ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

