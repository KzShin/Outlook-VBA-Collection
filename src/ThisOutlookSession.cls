Option Explicit

' ==============================================================================
' Module: ThisOutlookSession
' Description: Outlookイベントおよびマクロのエントリーポイント
' Dependencies: modLogger, modAutoFlag, modSendController, modMailSevenZip, modLaunchOWA
' ==============================================================================

' ------------------------------------------------------------------------------
' 1. イベントハンドラ (Events)
' ------------------------------------------------------------------------------

' 起動時の処理
Private Sub Application_Startup()
    ' 自動フラグ: 起動時に未読メールをチェック
    modAutoFlag.ProcessStartupUnread
End Sub

' メール受信時の処理
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    ' 自動フラグ: 新着メールをチェック
    modAutoFlag.ProcessNewMail EntryIDCollection
End Sub

' メール送信時の処理
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    ' 誤送信防止コントローラー: 送信前のチェックを実行
    modSendController.Execute Item, Cancel
End Sub

' ------------------------------------------------------------------------------
' 2. マクロ (Macros) - リボンやクイックアクセスツールバーに登録用
' ------------------------------------------------------------------------------

' 選択したメールの添付ファイルを保存・解凍するマクロ
Public Sub メールをダウンロードフォルダに保存()
    On Error GoTo EH

    ' RunID 生成 (yymmdd-hhnnss-SAVE)
    Dim runId As String
    runId = Format(Now, "yymmdd-hhnnss") & "-SAVE"

    ' ログ機能の初期化（IDセット）
    ' modMailSevenZipは、呼び出し元でIDがセットされていることを前提としています
    modLogger.SetRunId runId

    ' メイン処理の実行
    ' 変更: 英語名メソッドへの呼び出し
    modMailSevenZip.SaveAndExtractAttachments

    ' 後片付け
    modLogger.SetRunId "NoID"
    Exit Sub

EH:
    modLogger.Log "ThisOutlookSession", "ERROR: " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub

' WEB版Outlook (OWA) を開くマクロ
Public Sub WEB版Outlook起動()
    ' modLaunchOWA内でRunID生成・ログ出力まで完結しているため、直接呼び出します
    modLaunchOWA.LaunchOWA
End Sub