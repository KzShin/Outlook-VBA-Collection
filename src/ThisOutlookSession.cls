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
    ' modMailSevenZip内でRunID生成・ログ出力まで完結するよう変更したため、直接呼び出します
    modMailSevenZip.SaveAndExtractAttachments
End Sub

' WEB版Outlook (OWA) を開くマクロ
Public Sub WEB版Outlook起動()
    ' modLaunchOWA内でRunID生成・ログ出力まで完結しているため、直接呼び出します
    modLaunchOWA.LaunchOWA
End Sub

' 選択したメールをPDFとして印刷するマクロ
Public Sub 選択メールをPDF化()
    ' modMailToPDF内でRunID生成・ログ出力まで完結するため、直接呼び出します
    modMailToPDF.PrintMailToPDF
End Sub
