Attribute VB_Name = "modMailSevenZip"
Option Explicit

' ==============================================================================
' Module: modMailSevenZip
' Description: メール添付ファイルを保存し、必要に応じて7-Zipで解凍・結合を行う統合モジュール
'              - 単一メール選択時: 通常の添付ファイル保存とZip/7z解凍
'              - 複数メール選択時: 分割Zip(.001など)の保存、結合、解凍
' Dependencies: modLogger, Scripting.FileSystemObject, ADODB.Stream, WScript.Shell, MSXML2.DOMDocument
' Configuration: %APPDATA%\OutlookVBA\SevenZipPasswords.enc (暗号化パスワードリスト・推奨)
'                %APPDATA%\OutlookVBA\SevenZipPasswords.txt (平文パスワードリスト・下位互換用)
' ==============================================================================

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private m_FSO As Object

Private Function GetFSO() As Object
    If m_FSO Is Nothing Then Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set GetFSO = m_FSO
End Function

' ==============================================================================
' [Private] ログ出力ヘルパー（modLoggerへの委譲）
' ==============================================================================
Private Sub Log(ByVal msg As String)
    modLogger.Log "MailSevenZip", msg
End Sub

' ==============================================================================
' [Main] メイン処理 (エントリーポイント)
' ==============================================================================

' 選択したメールを保存し、アーカイブであれば解凍を試みる（単一/複数 自動判定）
Public Sub SaveAndExtractAttachments()
    On Error GoTo EH

    Dim runId As String
    runId = Format(Now, "yymmdd-hhnnss") & "-SAVE"
    modLogger.SetRunId runId

    Log "=== START メール保存・解凍処理 ==="

    Dim sel As Outlook.Selection
    Set sel = Application.ActiveExplorer.Selection

    If sel.Count = 0 Then
        MsgBox "メールを選択してください。", vbExclamation
        Log "選択なし：処理終了"
        modLogger.SetRunId "NoID"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If sel.Count = 1 Then
        Log "モード: 単一メール処理"
        ProcessSingleMail sel(1), fso
    Else
        Log "モード: 複数メール処理（分割Zip結合）"
        ProcessSplitArchives sel, fso
    End If

    Log "=== END メール保存・解凍処理 ==="
    modLogger.SetRunId "NoID"
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Number & vbCrLf & Err.Description, vbCritical
    modLogger.SetRunId "NoID"
End Sub


' ==============================================================================
' [Processor] 個別処理ロジック
' ==============================================================================

' --- 単一メール処理 ---
Private Sub ProcessSingleMail(ByVal objItem As Object, ByVal fso As Object)
    If objItem.Class <> olMail Then
        MsgBox "選択されたアイテムはメールではありません。", vbExclamation
        Log "非メールアイテム：処理終了"
        Exit Sub
    End If

    Dim mail As Outlook.MailItem
    Set mail = objItem
    Log "対象：" & mail.Subject

    ' ルートフォルダ構築
    Dim rootPath As String
    rootPath = Environ$("USERPROFILE") & "\Downloads\" & _
               Format(mail.ReceivedTime, "yymmdd_hhnnss") & "_" & SafeName(mail.Subject) & "\"

    If fso.FolderExists(rootPath) Then
        MsgBox "保存先フォルダが既に存在します。" & vbCrLf & rootPath, vbInformation
        Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
        Exit Sub
    End If

    fso.CreateFolder rootPath
    Log "フォルダ作成：" & rootPath

    ' 7-Zip 環境確認
    Dim sevenZipPath As String: sevenZipPath = modLogger.GetSevenZipPath()
    Dim has7zip As Boolean: has7zip = (Len(sevenZipPath) > 0)

    ' 添付ファイルの保存と解凍
    Dim att As Outlook.Attachment
    Dim savePath As String
    Dim outDir As String
    Dim cancelled As Boolean

    For Each att In mail.Attachments
        savePath = MakeUniqueFilePath(fso, fso.BuildPath(rootPath, att.FileName))

        Log "添付保存開始：" & att.FileName
        att.SaveAsFile savePath

        ' アーカイブ処理
        If has7zip And IsArchiveTarget(att.FileName) Then
            outDir = MakeUniqueFolderPath(fso, fso.BuildPath(rootPath, SafeName(fso.GetBaseName(att.FileName))))
            fso.CreateFolder outDir

            cancelled = False
            If TestThenExtractArchive(savePath, outDir, sevenZipPath, mail.ReceivedTime, 120, cancelled) Then
                Log "展開成功：" & outDir
            Else
                DeleteFolderIfEmpty fso, outDir
                If Not cancelled Then
                    MsgBox "パスワード候補では解凍できませんでした: " & att.FileName, vbInformation
                End If
            End If
        End If
    Next att

    ' 本文保存
    WriteTextUtf8 fso.BuildPath(rootPath, "メール本文.txt"), BuildMailInfoText(mail)
    Log "エクスプローラ起動"
    Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
End Sub

' --- 複数メール処理（分割Zip） ---
Private Sub ProcessSplitArchives(ByVal sel As Outlook.Selection, ByVal fso As Object)
    Dim targetItem As Object, mail As Outlook.MailItem
    Dim masterMail As Outlook.MailItem
    Dim att As Outlook.Attachment
    Dim baseName As String

    ' 1. マスターメール（.001）特定
    For Each targetItem In sel
        If TypeName(targetItem) = "MailItem" Then
            Set mail = targetItem
            For Each att In mail.Attachments
                If LCase$(Right$(att.FileName, 4)) = ".001" Then
                    Set masterMail = mail
                    baseName = fso.GetBaseName(att.FileName) ' 例: sample.zip
                    baseName = fso.GetBaseName(baseName)     ' 例: sample
                    Exit For
                End If
            Next
        End If
        If Not masterMail Is Nothing Then Exit For
    Next

    If masterMail Is Nothing Then
        MsgBox "選択されたメールの中に分割ファイル ('.001') が見つかりません。", vbExclamation
        Log "マスターメール未検出"
        Exit Sub
    End If

    ' 2. ルートフォルダ構築
    Dim rootPath As String
    rootPath = Environ$("USERPROFILE") & "\Downloads\" & _
               Format(masterMail.ReceivedTime, "yymmdd_hhnnss") & "_" & SafeName(masterMail.Subject) & "\"

    If fso.FolderExists(rootPath) Then
        MsgBox "保存先フォルダが既に存在します。" & vbCrLf & rootPath, vbInformation
        Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
        Exit Sub
    End If

    fso.CreateFolder rootPath
    Log "フォルダ作成：" & rootPath

    ' 3. 全メールの保存
    Dim firstPartPath As String
    For Each targetItem In sel
        If TypeName(targetItem) = "MailItem" Then
            Set mail = targetItem
            For Each att In mail.Attachments
                Dim savePath As String
                savePath = MakeUniqueFilePath(fso, fso.BuildPath(rootPath, att.FileName))

                att.SaveAsFile savePath
                Log "添付保存：" & att.FileName

                If LCase$(Right$(att.FileName, 4)) = ".001" Then firstPartPath = savePath

                Dim ext As String: ext = fso.GetExtensionName(att.FileName)
                If Len(ext) > 0 Then
                    WriteTextUtf8 fso.BuildPath(rootPath, "メール本文" & ext & ".txt"), BuildMailInfoText(mail)
                End If
            Next
        End If
    Next

    ' 4. 解凍処理
    Dim sevenZipPath As String: sevenZipPath = modLogger.GetSevenZipPath()
    If Len(sevenZipPath) > 0 And Len(firstPartPath) > 0 Then
        Dim outDir As String
        outDir = MakeUniqueFolderPath(fso, fso.BuildPath(rootPath, SafeName(baseName)))
        fso.CreateFolder outDir

        Dim cancelled As Boolean
        If TestThenExtractArchive(firstPartPath, outDir, sevenZipPath, masterMail.ReceivedTime, 120, cancelled) Then
            Log "展開成功：" & outDir
        Else
            DeleteFolderIfEmpty fso, outDir
            If Not cancelled Then
                MsgBox "パスワード候補では解凍できませんでした: " & fso.GetFileName(firstPartPath), vbInformation
            End If
        End If
    End If

    Log "エクスプローラ起動"
    Shell "explorer.exe " & """" & rootPath & """", vbNormalFocus
End Sub


' ==============================================================================
' [Core Logic] 解凍フロー制御
' ==============================================================================

Private Function IsArchiveTarget(ByVal fileName As String) As Boolean
    Dim dotPos As Long: dotPos = InStrRev(fileName, ".")
    If dotPos = 0 Then Exit Function
    Dim ext As String: ext = LCase$(Mid$(fileName, dotPos + 1))
    IsArchiveTarget = (ext = "zip" Or ext = "7z")
End Function

Private Function TestThenExtractArchive(ByVal zipPath As String, ByVal outDir As String, _
                                        ByVal sevenZipPath As String, ByVal receivedTime As Date, _
                                        ByVal timeoutSec As Long, ByRef userCancelled As Boolean) As Boolean
    userCancelled = False
    Dim rc As Long

    Log "テスト（パス無し）実行"
    rc = SevenZipTest(zipPath, sevenZipPath, "", timeoutSec)
    If rc <= 1 Then
        rc = SevenZipExtract(zipPath, outDir, sevenZipPath, "", timeoutSec)
        PauseSeconds 0.5
        If rc = 0 Or HasNonZeroFileDeep(outDir) Then
            TestThenExtractArchive = True
            Exit Function
        End If
    End If

    Dim cands As Collection: Set cands = GetPasswordCandidates(receivedTime)
    Dim pw As Variant, idx As Long: idx = 0
    For Each pw In cands
        idx = idx + 1
        rc = SevenZipTest(zipPath, sevenZipPath, CStr(pw), timeoutSec)
        If rc <= 1 Then
            rc = SevenZipExtract(zipPath, outDir, sevenZipPath, CStr(pw), timeoutSec)
            PauseSeconds 0.5
            If rc = 0 Or HasNonZeroFileDeep(outDir) Then
                TestThenExtractArchive = True
                Exit Function
            End If
        End If
    Next pw

    If PromptAndTryPassword(zipPath, outDir, sevenZipPath, timeoutSec, userCancelled) Then
        TestThenExtractArchive = True
        Exit Function
    End If
    TestThenExtractArchive = False
End Function

Private Function PromptAndTryPassword(ByVal zipPath As String, ByVal outDir As String, _
                                      ByVal sevenZipPath As String, ByVal timeoutSec As Long, _
                                      ByRef userCancelled As Boolean) As Boolean
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pw As String, rc As Long
    userCancelled = False

    Do
        pw = InputBox("登録済みパスワードでは解凍できませんでした。" & vbCrLf & _
                      "解凍用パスワードを入力してください（キャンセルで中止）。", "パスワード入力")
        If Len(pw) = 0 Then
            userCancelled = True
            PromptAndTryPassword = False
            Exit Function
        End If

        On Error Resume Next
        If fso.FolderExists(outDir) Then fso.DeleteFolder outDir, True
        fso.CreateFolder outDir
        On Error GoTo EH

        rc = SevenZipTest(zipPath, sevenZipPath, pw, timeoutSec)
        If rc <= 1 Then
            rc = SevenZipExtract(zipPath, outDir, sevenZipPath, pw, timeoutSec)
            PauseSeconds 0.5
            If rc = 0 Or HasNonZeroFileDeep(outDir) Then
                PromptAndTryPassword = True
                Exit Function
            End If
        End If
    Loop
EH:
    PromptAndTryPassword = False
End Function


' ==============================================================================
' [7-Zip Wrapper] コマンドライン実行
' ==============================================================================

Private Function SevenZipTest(ByVal zipPath As String, ByVal sevenZipPath As String, _
                              ByVal password As String, ByVal timeoutSec As Long) As Long
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim baseCmd As String
    baseCmd = """" & sevenZipPath & """ t -y """ & zipPath & """ -bso0 -bse0 -bsp0"
    ' 空パスワードであっても常に -p を付与し、対話型プロンプト待ちを防止する
    Dim cmd As String: cmd = baseCmd & " -p""" & password & """"
    
    If Len(password) > 0 Then
        Log "7zテスト：" & baseCmd & " -p""" & MaskPassword(password) & """"
    Else
        Log "7zテスト：" & baseCmd & " -p"""""
    End If

    Dim proc As Object: Set proc = sh.Exec(cmd)
    On Error Resume Next
    proc.StdIn.Close ' 万が一の対話待ちを二重防止
    On Error GoTo 0
    SevenZipTest = WaitProcessWithTimeout(proc, timeoutSec)
    Set proc = Nothing
    Set sh = Nothing
End Function

Private Function SevenZipExtract(ByVal zipPath As String, ByVal outDir As String, _
                                 ByVal sevenZipPath As String, ByVal password As String, _
                                 ByVal timeoutSec As Long) As Long
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim baseCmd As String
    baseCmd = """" & sevenZipPath & """ x -y """ & zipPath & """ -o""" & outDir & """ -bso0 -bse0 -bsp0"
    ' 空パスワードであっても常に -p を付与し、対話型プロンプト待ちを防止する
    Dim cmd As String: cmd = baseCmd & " -p""" & password & """"
    
    If Len(password) > 0 Then
        Log "7z抽出：" & baseCmd & " -p""" & MaskPassword(password) & """"
    Else
        Log "7z抽出：" & baseCmd & " -p"""""
    End If

    Dim proc As Object: Set proc = sh.Exec(cmd)
    On Error Resume Next
    proc.StdIn.Close ' 万が一の対話待ちを二重防止
    On Error GoTo 0
    SevenZipExtract = WaitProcessWithTimeout(proc, timeoutSec)
    Set proc = Nothing
    Set sh = Nothing
End Function


' ==============================================================================
' [Config] パスワード候補管理
' ==============================================================================

Private Function GetPasswordCandidates(ByVal receivedTime As Date) As Collection
    Dim col As New Collection
    Dim folderPath As String: folderPath = Environ$("APPDATA") & "\OutlookVBA"
    Dim encPath As String: encPath = folderPath & "\SevenZipPasswords.enc"
    Dim txtPath As String: txtPath = folderPath & "\SevenZipPasswords.txt"

    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    On Error GoTo 0

    ' 1. 暗号化ファイル (.enc) が存在する場合は最優先で復号して読み込み
    If fso.FileExists(encPath) Then
        If LoadPasswordsFromEncryptedFile(encPath, receivedTime, col) Then
            Log "PW候補読込完了(暗号化): " & col.Count & "件"
            Set GetPasswordCandidates = col
            Set fso = Nothing
            Exit Function
        Else
            Log "PW候補読込警告: 暗号化ファイルの復号に失敗しました。平文ファイルの確認へ進みます。"
        End If
    End If

    ' 2. 平文ファイル (.txt) の下位互換フォールバック
    If fso.FileExists(txtPath) Then
        If LoadPasswordsFromFile(txtPath, receivedTime, col) Then
            Log "PW候補読込注意: 平文ファイルを使用しています。Protect-SevenZipPassword.ps1 による暗号化を推奨します (" & col.Count & "件)"
            Set GetPasswordCandidates = col
            Set fso = Nothing
            Exit Function
        End If
    End If

    Set fso = Nothing
    Set GetPasswordCandidates = col
End Function

' 暗号化ファイル (.enc) を PowerShell (DPAPI / SecureString) で復号して読み込み
Private Function LoadPasswordsFromEncryptedFile(ByVal filePath As String, ByVal receivedTime As Date, ByRef outCol As Collection) As Boolean
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Set fso = Nothing

    Dim yyyy As String: yyyy = Format(receivedTime, "yyyy")
    Dim yy As String: yy = Right$(yyyy, 2)
    Dim mm As String: mm = Format(receivedTime, "mm")
    Dim dd As String: dd = Format(receivedTime, "dd")

    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    Dim psScript As String
    psScript = "Import-Module $PSHOME\Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1 -ErrorAction SilentlyContinue; " & _
               "& { param($p) if (Test-Path -LiteralPath $p) { " & _
               "Get-Content -LiteralPath $p | ForEach-Object { " & _
               "$l = $_.Trim(); if ($l.Length -gt 0) { " & _
               "try { $s = ConvertTo-SecureString $l; $b = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($s); " & _
               "try { [Console]::WriteLine([System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes([System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($b)))) } " & _
               "finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b) } " & _
               "} catch {} } } } } '" & filePath & "'"

    Dim fullCmd As String
    fullCmd = "powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command """ & psScript & """"

    Dim oExec As Object: Set oExec = wsh.Exec(fullCmd)
    Do While oExec.Status = 0
        DoEvents
    Loop

    Dim outText As String
    outText = oExec.StdOut.ReadAll()
    Set oExec = Nothing
    Set wsh = Nothing

    If Len(Trim$(outText)) = 0 Then
        LoadPasswordsFromEncryptedFile = False
        Exit Function
    End If

    Dim lines() As String
    lines = Split(Replace(outText, vbCrLf, vbLf), vbLf)

    Dim i As Long, b64Line As String, plain As String, cnt As Long
    For i = LBound(lines) To UBound(lines)
        b64Line = Trim$(lines(i))
        If Len(b64Line) > 0 Then
            plain = DecodeBase64Utf8(b64Line)
            If Len(plain) > 0 Then
                plain = ExpandPasswordTemplate(plain, yyyy, yy, mm, dd)
                outCol.Add plain
                cnt = cnt + 1
            End If
        End If
    Next i

    LoadPasswordsFromEncryptedFile = (cnt > 0)
    Exit Function

EH:
    LoadPasswordsFromEncryptedFile = False
End Function

' 平文テキストファイル (.txt) の読み込み（下位互換用）
Private Function LoadPasswordsFromFile(ByVal filePath As String, ByVal receivedTime As Date, ByRef outCol As Collection) As Boolean
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function

    Dim yyyy As String: yyyy = Format(receivedTime, "yyyy")
    Dim yy As String: yy = Right$(yyyy, 2)
    Dim mm As String: mm = Format(receivedTime, "mm")
    Dim dd As String: dd = Format(receivedTime, "dd")

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8": stm.Open: stm.LoadFromFile filePath
    Dim allText As String: allText = stm.ReadText(-1): stm.Close

    Dim lines() As String: lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
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
                expanded = ExpandPasswordTemplate(expanded, yyyy, yy, mm, dd)
                outCol.Add expanded
                cnt = cnt + 1
            End If
        End If
    Next i
    LoadPasswordsFromFile = (cnt > 0)
    Exit Function
EH:
    LoadPasswordsFromFile = False
End Function

' パスワード文字列の日付プレースホルダーを展開
Private Function ExpandPasswordTemplate(ByVal rawPw As String, ByVal yyyy As String, ByVal yy As String, ByVal mm As String, ByVal dd As String) As String
    Dim res As String: res = rawPw
    res = Replace(res, "{yyyy}", yyyy)
    res = Replace(res, "{yy}", yy)
    res = Replace(res, "{mm}", mm)
    res = Replace(res, "{dd}", dd)
    ExpandPasswordTemplate = res
End Function

' Base64文字列 (UTF-8) を平文文字列にデコード
Private Function DecodeBase64Utf8(ByVal b64Text As String) As String
    On Error GoTo EH
    Dim xmlDoc As Object: Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Dim el As Object: Set el = xmlDoc.createElement("b64")
    el.DataType = "bin.base64"
    el.Text = b64Text

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.Write el.nodeTypedValue
    stm.Position = 0
    stm.Type = 2 ' adTypeText
    stm.Charset = "UTF-8"
    DecodeBase64Utf8 = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
    Set el = Nothing
    Set xmlDoc = Nothing
    Exit Function
EH:
    DecodeBase64Utf8 = ""
End Function


' ==============================================================================
' [System] プロセス制御ユーティリティ
' ==============================================================================

Private Function WaitProcessWithTimeout(ByVal proc As Object, ByVal timeoutSeconds As Long) As Long
    Dim startTick As Double: startTick = Timer
    
    ' WshExecのStdOut.AtEndOfStreamはプロセス待機中に同期ブロックするため、
    ' proc.Status（0: 実行中, 1: 終了）で安全にポーリング待機を行う
    Do While proc.Status = 0
        If ElapsedSeconds(startTick) >= timeoutSeconds Then
            On Error Resume Next
            proc.Terminate
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
    IsProcessRunning = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub TerminateProcessByPID(ByVal pid As Long)
    On Error Resume Next
    Dim svc As Object: Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Dim obj As Object: Set obj = svc.Get("Win32_Process.Handle='" & pid & "'")
    If Not obj Is Nothing Then obj.Terminate
    Set obj = Nothing
    Set svc = Nothing
    On Error GoTo 0
End Sub

Private Function ElapsedSeconds(ByVal startTick As Double) As Double
    Dim t As Double: t = Timer
    If t >= startTick Then ElapsedSeconds = t - startTick Else ElapsedSeconds = (86400# - startTick) + t
End Function

Private Sub PauseSeconds(ByVal seconds As Double)
    Dim st As Double: st = Timer
    Do While ElapsedSeconds(st) < seconds
        Sleep 50
        DoEvents
    Loop
End Sub


' ==============================================================================
' [Utils] ファイル・テキスト操作ユーティリティ
' ==============================================================================

Private Function MaskPassword(ByVal pw As String) As String
    Dim n As Long: n = Len(pw)
    If n = 0 Then Exit Function
    If n <= 2 Then
        MaskPassword = String(n, "*")
    Else
        MaskPassword = Left$(pw, 1) & String(n - 1, "*")
    End If
End Function

Private Function HasNonZeroFileDeep(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Dim fso As Object: Set fso = GetFSO()
    If Not fso.FolderExists(folderPath) Then Exit Function
    Dim fld As Object: Set fld = fso.GetFolder(folderPath)

    Dim f As Object
    For Each f In fld.Files
        If f.Size > 0 Then HasNonZeroFileDeep = True: Exit Function
    Next f

    Dim subf As Object
    For Each subf In fld.SubFolders
        If HasNonZeroFileDeep(subf.Path) Then HasNonZeroFileDeep = True: Exit Function
    Next subf
    On Error GoTo 0
End Function

Private Sub DeleteFolderIfEmpty(ByVal fso As Object, ByVal folderPath As String)
    On Error Resume Next
    If fso.FolderExists(folderPath) Then
        Dim fld As Object: Set fld = fso.GetFolder(folderPath)
        If fld.Files.Count = 0 And fld.SubFolders.Count = 0 Then fso.DeleteFolder folderPath, True
    End If
End Sub

Private Function SafeName(ByVal s As String) As String
    Dim r As String: r = s
    r = Replace(r, "<", "_"): r = Replace(r, ">", "_"): r = Replace(r, ":", "_")
    r = Replace(r, """", "_"): r = Replace(r, "/", "_"): r = Replace(r, "\", "_")
    r = Replace(r, "|", "_"): r = Replace(r, "?", "_"): r = Replace(r, "*", "_")
    r = Replace(r, "%", "_"): r = Replace(r, "&", "_"): r = Replace(r, "^", "_")
    r = Replace(r, vbCr, "_"): r = Replace(r, vbLf, "_"): r = Replace(r, vbTab, "_")
    If Len(r) > 150 Then r = Left$(r, 150)
    SafeName = Trim$(r)
End Function

Private Function MakeUniqueFilePath(ByVal fso As Object, ByVal path As String) As String
    If Not fso.FileExists(path) Then MakeUniqueFilePath = path: Exit Function
    Dim cand As String, i As Long: i = 2
    Do
        cand = fso.BuildPath(fso.GetParentFolderName(path), fso.GetBaseName(path) & " (" & i & ")." & fso.GetExtensionName(path))
        If Not fso.FileExists(cand) Then MakeUniqueFilePath = cand: Exit Function
        i = i + 1
    Loop
End Function

Private Function MakeUniqueFolderPath(ByVal fso As Object, ByVal path As String) As String
    If Not fso.FolderExists(path) Then MakeUniqueFolderPath = path: Exit Function
    Dim cand As String, i As Long: i = 2
    Do
        cand = fso.BuildPath(fso.GetParentFolderName(path), fso.GetFileName(path) & " (" & i & ")")
        If Not fso.FolderExists(cand) Then MakeUniqueFolderPath = cand: Exit Function
        i = i + 1
    Loop
End Function

Private Function BuildMailInfoText(ByVal mail As Outlook.MailItem) As String
    Dim sb As String
    sb = "受信日時: " & mail.ReceivedTime & vbCrLf & _
         "From: " & mail.SenderName & " <" & mail.SenderEmailAddress & ">" & vbCrLf & _
         "To: " & mail.To & vbCrLf & _
         "CC: " & mail.CC & vbCrLf & _
         "件名: " & mail.Subject & vbCrLf & vbCrLf & _
         "メール本文:" & vbCrLf & mail.Body & vbCrLf
    BuildMailInfoText = sb
End Function

Private Sub WriteTextUtf8(ByVal path As String, ByVal text As String)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8": stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2
    stm.Close
    Set stm = Nothing
End Sub
