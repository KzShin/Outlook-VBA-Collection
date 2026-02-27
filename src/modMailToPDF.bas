Attribute VB_Name = "modMailToPDF"
Option Explicit

' ==============================================================================
' Module: modMailToPDF
' Description: 選択したメールをPDFプリンタ経由でPDF化するモジュール
' Dependencies: WScript.Network, WMI, ADODB.Stream, modLogger
' Configuration: %APPDATA%\OutlookVBA\config.ini ([MailToPDF] Section)
' ==============================================================================

' --- グローバル変数 (Module Level) ---
Private g_RunId As String
Private g_ConfigCache As Object
Private g_IsConfigLoaded As Boolean

' ==============================================================================
' [Public] 公開インターフェース・初期化
' ==============================================================================

Public Sub PrintMailToPDF()
    On Error GoTo EH
    
    ' 実行ID生成 (yymmdd-hhnnss-TOPDF)
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-TOPDF"
    
    ' 自モジュールと共通ロガーの両方にIDを設定
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START メールPDF化処理 ==="
    
    ' メイン処理の呼び出し
    ProcessPrintToPDF
    
    Log "=== END メールPDF化処理 ==="
    modLogger.SetRunId "NoID"
    Exit Sub

EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "PDF化処理中にエラーが発生しました:" & vbCrLf & Err.Description, vbCritical, "MailToPDF"
    modLogger.SetRunId "NoID"
End Sub

' ------------------------------------------------------------------------------
' [Private] ログ・ID管理ヘルパー
' ------------------------------------------------------------------------------
Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

' 共通ロガーへの委譲
Private Sub Log(ByVal msg As String)
    modLogger.Log "MailToPDF", msg
End Sub


' ==============================================================================
' [Main] メイン処理
' ==============================================================================

Private Sub ProcessPrintToPDF()
    On Error GoTo EH
    
    ' 1. 対象アイテムの確認
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "メールを選択してください。", vbExclamation, "MailToPDF"
        Log "選択アイテムなし"
        Exit Sub
    End If
    
    Dim objItem As Object
    Set objItem = Application.ActiveExplorer.Selection(1)
    
    If objItem.Class <> olMail Then
        MsgBox "選択されたアイテムはメールではありません。", vbExclamation, "MailToPDF"
        Log "対象外アイテム: " & TypeName(objItem)
        Exit Sub
    End If
    
    Dim mailItem As Outlook.MailItem
    Set mailItem = objItem
    Log "対象メール: " & mailItem.Subject
    
    ' 2. 設定の読み込み
    Dim pdfPrinter As String
    Dim physicalPrinter As String
    
    pdfPrinter = GetConfigValue("PdfPrinterName", "Microsoft Print to PDF")
    physicalPrinter = GetConfigValue("PhysicalPrinterName", "Auto")
    
    ' 3. 元のプリンタの自動判別機能
    If physicalPrinter = "" Or LCase$(physicalPrinter) = "auto" Then
        physicalPrinter = GetCurrentDefaultPrinter()
        Log "元プリンタを自動取得しました: " & physicalPrinter
    Else
        Log "元プリンタを設定ファイルから取得しました: " & physicalPrinter
    End If
    
    If physicalPrinter = "" Then
        Log "元のプリンタ名が取得/設定されていません"
        MsgBox "元のプリンタ名が取得できませんでした。処理を中止します。", vbExclamation, "MailToPDF"
        Exit Sub
    End If
    
    ' 4. プリンタ変更と印刷処理
    Dim wshNet As Object
    Set wshNet = CreateObject("WScript.Network")
    
    ' --- PDFプリンタに変更 ---
    Log "デフォルトプリンタを変更: " & pdfPrinter
    On Error Resume Next
    wshNet.SetDefaultPrinter pdfPrinter
    If Err.Number <> 0 Then
        Log "PDFプリンタへの変更失敗: " & Err.Description
        MsgBox "PDFプリンタ (" & pdfPrinter & ") の設定に失敗しました。", vbCritical, "MailToPDF"
        On Error GoTo EH
        Exit Sub
    End If
    On Error GoTo EH
    
    ' --- 印刷実行 ---
    Log "印刷開始 (PrintOut)"
    mailItem.PrintOut
    Log "印刷コマンド送信完了"
    
    ' --- 元のプリンタに戻す ---
    Log "デフォルトプリンタを復元: " & physicalPrinter
    On Error Resume Next
    wshNet.SetDefaultPrinter physicalPrinter
    If Err.Number <> 0 Then
        Log "元のプリンタへの復元失敗: " & Err.Description
        MsgBox "元のプリンタ (" & physicalPrinter & ") に戻す際にエラーが発生しました。", vbCritical, "MailToPDF"
    End If
    On Error GoTo EH
    
    Exit Sub

EH:
    ' 上位プロシージャへエラーを引き継ぐ
    Err.Raise Err.Number, , Err.Description
End Sub


' ==============================================================================
' [Config] 設定読み込み
' ==============================================================================

Private Sub LoadConfig()
    On Error GoTo EH
    Set g_ConfigCache = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim configPath As String
    configPath = Environ$("APPDATA") & "\OutlookVBA\config.ini"
    
    If Not fso.FileExists(configPath) Then
        Log "Config file not found: " & configPath
        g_IsConfigLoaded = True
        Exit Sub
    End If
    
    ' ADODB.StreamによるUTF-8読み込み
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
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
            
            ' セクション判定
            If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
            
            ' [MailToPDF] セクションのみ読み込む
            ElseIf currentSection = "mailtopdf" Then
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
        Log "Config Key Not Found: [" & key & "] -> Using Default: " & defaultVal
        GetConfigValue = defaultVal
    End If
End Function


' ==============================================================================
' [Logic] ユーティリティ
' ==============================================================================

' WMIを使用して現在設定されているデフォルトプリンタの名前を取得する
Private Function GetCurrentDefaultPrinter() As String
    On Error Resume Next
    Dim wmi As Object
    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Dim colPrinters As Object
    ' Default が True に設定されているプリンタを検索
    Set colPrinters = wmi.ExecQuery("Select * from Win32_Printer Where Default = True")
    
    Dim printer As Object
    For Each printer In colPrinters
        GetCurrentDefaultPrinter = printer.Name
        Exit For
    Next
    On Error GoTo 0
End Function
