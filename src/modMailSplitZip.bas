Attribute VB_Name = "modMailSplitZip"
Option Explicit

' ==============================================================================
' Module: modMailSplitZip
' Description: 複数メールに分割されたZipファイル(.001, .002...)を一括保存し結合・解凍する
' Dependencies: modLogger, Scripting.FileSystemObject, ADODB.Stream, WScript.Shell
' Configuration: %APPDATA%\OutlookVBA\SevenZipPasswords.txt (解凍パスワードリスト)
' ==============================================================================

' --- ThisOutlookSessionでの呼び出し例 ---
' Public Sub 選択メールの分割Zipを結合解凍()
'     modMailSplitZip.SaveAndExtractSplitArchives
' End Sub

' ==============================================================================
' [Private] ログ出力ヘルパー（modLoggerへの委譲）
' ==============================================================================

Private Sub Log(ByVal msg As String)
    modLogger.Log "MailSplitZip", msg
End Sub


' ==============================================================================
' [Main] メイン処理
' ==============================================================================

' 選択した複数メールから分割ファイルを保存し、結合・解凍を行う
Public Sub SaveAndExtractSplitArchives()
    On Error GoTo EH

    ' 実行ID生成とセット
    Dim runId As String
    runId = Format(Now, "yymmdd-hhnnss") & "-SPLIT"
    modLogger.SetRunId runId

    Log "=== START 分割Zip結合解凍処理 ==="

    ' 1. メール選択チェック
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "分割ファイルが添付されたメールを複数選択してください。", vbExclamation
        Log "選択なし：処理終了"
        modLogger.SetRunId "NoID"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 2. マスターメール（.001を持つメール）の特定
    Dim targetItem As Object
    Dim mail As Outlook.MailItem
    Dim masterMail As Outlook.MailItem
    Dim att As Outlook.Attachment
    Dim baseName As String

    For Each targetItem In Application.ActiveExplorer.Selection
        If TypeName(targetItem) = "MailItem" Then
            Set mail = targetItem
            For Each att In mail.Attachments
                If LCase$(Right$(att.FileName, 4)) = ".001" Then
                    Set masterMail = mail
                    ' 拡張子を除いたベース名を取得 (例: sample.zip.001 -> sample.zip)
                    baseName = fso.GetBaseName(att.FileName)
                    ' さらに ".zip" や ".7z" があれば除外してフォルダ名用にする
                    baseName = fso.GetBaseName(baseName)
                    Exit For
                End If
            Next
        End If
        If Not masterMail Is Nothing Then Exit For
    Next

    If masterMail Is Nothing Then
        MsgBox "選択されたメールの中に '.001' の添付ファイルが見つかりません。", vbExclamation
        Log "マスターメール(.001)未検出により終了"
        modLogger.SetRunId "NoID"
        Exit Sub
    End If

    Log "マスターメール特定：" & masterMail.Subject & " (BaseName: " & baseName & ")"

    ' 3. 保存先ルートパスの構築
    Dim receivedDate As String
    receivedDate = Format(masterMail.ReceivedTime, "yymmdd_hhnnss")

    Dim safeSubj As String
    safeSubj = SafeName(masterMail.Subject)

    Dim rootPath As String
    rootPath = Environ$("USERPROFILE") & "\Downloads\" & receivedDate & "_" & safeSubj & "\"
    Log "保存先ルート：" & rootPath

    If fso.FolderExists(rootPath) Then
        MsgBox "保存先フォルダが既に存在します。既存フォルダを開きます。" & vbCrLf & rootPath, vbInformation
        Log "既存フォルダ検出：処理中止 " & rootPath
        Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
        modLogger.SetRunId "NoID"
        Exit Sub
    End If

    fso.CreateFolder rootPath
    Log "フォルダ作成済み：" & rootPath

    ' 4. 各メールの添付ファイルと本文の保存
    Dim firstPartPath As String

    For Each targetItem In Application.ActiveExplorer.Selection
        If TypeName(targetItem) = "MailItem" Then
            Set mail = targetItem
            Log "メール処理：" & mail.Subject

            For Each att In mail.Attachments
                Dim savePath As String
                savePath = fso.BuildPath(rootPath, att.FileName)
                savePath = MakeUniqueFilePath(fso, savePath)

                ' 添付ファイルの保存
                Log "添付保存開始：""" & att.FileName & """ -> " & savePath
                att.SaveAsFile savePath
                Log "添付保存完了：" & savePath

                ' .001ファイルのフルパスを保持（解凍の起点とするため）
                If LCase$(Right$(att.FileName, 4)) = ".001" Then
                    firstPartPath = savePath
                End If

                ' 拡張子を取得して本文テキストを保存（例: 001, 002）
                Dim ext As String
                ext = fso.GetExtensionName(att.FileName)

                If Len(ext) > 0 Then
                    Dim infoPath As String
                    infoPath = fso.BuildPath(rootPath, "メール本文" & ext & ".txt")
                    Log "本文書き出し開始：" & infoPath
                    WriteTextUtf8 infoPath, BuildMailInfoText(mail)
                    Log "本文書き出し完了：" & infoPath
                End If
            Next
        End If
    Next

    ' 5. 7-Zip 環境確認と解凍
    Dim sevenZipPath As String
    sevenZipPath = modLogger.GetSevenZipPath()
    Dim has7zip As Boolean
    has7zip = (Len(sevenZipPath) > 0)

    If Not has7zip Then
        MsgBox "7-Zip が見つかりません。ファイルの保存のみ完了しました。", vbExclamation
        Log "7-Zip未検出：展開スキップ"
        GoTo FinishProcess
    End If

    If Len(firstPartPath) = 0 Then
        Log ".001ファイルのパスが特定できないため解凍スキップ"
        GoTo FinishProcess
    End If

    ' 展開先フォルダ作成
    Dim outDir As String
    outDir = fso.BuildPath(rootPath, SafeName(baseName))
    outDir = MakeUniqueFolderPath(fso, outDir)
    fso.CreateFolder outDir
    Log "展開先フォルダ作成：" & outDir

    ' 結合解凍実行
    Dim cancelled As Boolean
    Dim ok As Boolean
    Log "事前テスト開始：" & firstPartPath

    ok = TestThenExtractArchive(firstPartPath, outDir, sevenZipPath, masterMail.ReceivedTime, 120, cancelled)

    If ok Then
        Log "展開成功：" & outDir
    Else
        Log "展開失敗：全候補不一致またはタイムアウト。後処理を実行"
        DeleteFolderIfEmpty fso, outDir

        If cancelled Then
            Log "ユーザーキャンセルにより終了"
        Else
            MsgBox "パスワード候補では解凍できませんでした: " & fso.GetFileName(firstPartPath) & vbCrLf & _
                   "分割ファイルは保存したままにしています。", vbInformation
        End If
    End If

FinishProcess:
    ' 6. 完了後のフォルダ表示
    Log "エクスプローラ起動：" & rootPath
    Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus

    Log "=== END 分割Zip結合解凍処理 ==="

    modLogger.SetRunId "NoID"
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Number & vbCrLf & Err.Description, vbCritical
    modLogger.SetRunId "NoID"
End Sub


' ==============================================================================
' [Core Logic] 解凍フロー制御 (modMailSevenZip 準拠)
' ==============================================================================

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

' 手入力による解凍試行
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

Private Function SevenZipTest(ByVal zipPath As String, ByVal sevenZipPath As String, _
                              ByVal password As String, ByVal timeoutSeconds As Long) As Long
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim baseCmd As String
    baseCmd = """" & sevenZipPath & """ t -y " & _
              """" & zipPath & """" & _
              " -bso0 -bse0 -bsp0"

    Dim cmd As String
    Dim logCmd As String

    cmd = baseCmd & " -p""" & password & """"
    logCmd = baseCmd & " -p""" & MaskPassword(password) & """"

    Log "7zテスト起動：" & logCmd
    Dim proc As Object
    Set proc = sh.Exec(cmd)

    SevenZipTest = WaitProcessWithTimeout(proc, timeoutSeconds)
    Log "7zテスト終了コード：" & SevenZipTest
End Function

Private Function SevenZipExtract(ByVal zipPath As String, ByVal outDir As String, _
                                 ByVal sevenZipPath As String, ByVal password As String, _
                                 ByVal timeoutSeconds As Long) As Long
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")

    Dim baseCmd As String
    baseCmd = """" & sevenZipPath & """ x -y " & _
              """" & zipPath & """ -o""" & outDir & """" & _
              " -bso0 -bse0 -bsp0"

    Dim cmd As String
    Dim logCmd As String

    cmd = baseCmd
    logCmd = baseCmd

    If Len(password) > 0 Then
        cmd = cmd & " -p""" & password & """"
        logCmd = logCmd & " -p""" & MaskPassword(password) & """"
    End If

    Log "7z抽出起動：" & logCmd
    Dim proc As Object
    Set proc = sh.Exec(cmd)

    SevenZipExtract = WaitProcessWithTimeout(proc, timeoutSeconds)
    Log "7z抽出終了コード：" & SevenZipExtract
End Function


' ==============================================================================
' [Config] パスワード候補管理
' ==============================================================================

Private Function GetPasswordCandidates(ByVal receivedTime As Date) As Collection
    Dim col As New Collection
    Dim appData As String
    appData = Environ$("APPDATA")

    Dim folderPath As String
    folderPath = appData & "\OutlookVBA"

    Dim listPath As String
    listPath = folderPath & "\SevenZipPasswords.txt"

    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
        Log "設定フォルダ作成: " & folderPath
    End If
    On Error GoTo 0

    Dim loaded As Boolean
    loaded = LoadPasswordsFromFile(listPath, receivedTime, col)

    If Not loaded Then
        Log "パスワードリスト読込なし（0件またはファイル未存在）: " & listPath
    End If

    Set GetPasswordCandidates = col
End Function

Private Function LoadPasswordsFromFile(ByVal filePath As String, _
                                       ByVal receivedTime As Date, _
                                       ByRef outCol As Collection) As Boolean
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        LoadPasswordsFromFile = False
        Exit Function
    End If

    Dim yyyy As String, yy As String, mm As String, dd As String
    yyyy = Format(receivedTime, "yyyy")
    yy = Right$(yyyy, 2)
    mm = Format(receivedTime, "mm")
    dd = Format(receivedTime, "dd")

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
    Dim isTarget As Boolean
    For i = LBound(lines) To UBound(lines)
        raw = Trim$(lines(i))
        isTarget = False

        If Len(raw) > 0 Then
            If Len(raw) >= 2 And Left$(raw, 1) = """" And Right$(raw, 1) = """" Then
                ' ダブルクォート囲み：クォートを除去して採用（"//" 始まりや空白保持も可能）
                expanded = Mid$(raw, 2, Len(raw) - 2)
                isTarget = True
            ElseIf Left$(raw, 2) <> "//" Then
                ' 通常のパスワード行（// コメント行以外、# 始まりもそのまま許可）
                expanded = raw
                isTarget = True
            End If

            If isTarget Then
                expanded = Replace(expanded, "{yyyy}", yyyy)
                expanded = Replace(expanded, "{yy}", yy)
                expanded = Replace(expanded, "{mm}", mm)
                expanded = Replace(expanded, "{dd}", dd)

                outCol.Add expanded
                cnt = cnt + 1
                Log "候補追加：" & MaskPassword(expanded)
            End If
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

Private Function WaitProcessWithTimeout(ByVal proc As Object, ByVal timeoutSeconds As Long) As Long
    Dim startTick As Single
    startTick = Timer

    Do
        Do While Not proc.StdOut.AtEndOfStream
            Dim s As String: s = proc.StdOut.Read(1024)
        Loop
        Do While Not proc.StdErr.AtEndOfStream
            Dim e As String: e = proc.StdErr.Read(1024)
        Loop

        If Not IsProcessRunning(proc) Then Exit Do

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

Private Function MaskPassword(ByVal pw As String) As String
    If Len(pw) = 0 Then
        MaskPassword = ""
        Exit Function
    End If

    Dim i As Long
    Dim res As String
    res = ""

    For i = 1 To Len(pw)
        If i Mod 2 = 1 Then
            res = res & Mid$(pw, i, 1)
        Else
            res = res & "*"
        End If
    Next i

    MaskPassword = res
End Function

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

Private Sub WriteTextUtf8(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub