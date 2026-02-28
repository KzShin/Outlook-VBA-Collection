Option Explicit

' ==============================================================================
' Module: ThisOutlookSession
' Description: Outlookイベントおよびマクロのエントリーポイント
' Dependencies: modLogger, modAutoFlag, modSendController, modMailSevenZip, modLaunchOWA, modMailToPDF, modMailOpen
' ==============================================================================

' --- メール開封制御用のイベント変数 ---
Private WithEvents myExplorer As Outlook.Explorer
Private WithEvents myItem As Outlook.MailItem

' ------------------------------------------------------------------------------
' 1. イベントハンドラ (Events)
' ------------------------------------------------------------------------------

' 起動時の処理
Private Sub Application_Startup()
    ' 自動フラグ: 起動時に未読メールをチェック
    modAutoFlag.ProcessStartupUnread
    
    ' メール開封制御: Explorerを取得し、初期選択アイテムを紐付け
    Set myExplorer = Application.ActiveExplorer
    If Not myExplorer Is Nothing Then
        Call myExplorer_SelectionChange
    End If
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

' ==============================================================================
' [MailOpen Control] メール開封のフック処理
' ==============================================================================

' 選択が変わるたびに、現在選択されているメールをmyItemに紐付ける
' （これにより、誤爆クリックしたアイテムのOpenイベントを確実にフックできます）
Private Sub myExplorer_SelectionChange()
    On Error Resume Next
    Set myItem = Nothing
    
    If myExplorer.Selection.Count = 1 Then
        If myExplorer.Selection(1).Class = olMail Then
            Set myItem = myExplorer.Selection(1)
        End If
    End If
End Sub

' メールのポップアップ（Openイベント）をフック
Private Sub myItem_Open(Cancel As Boolean)
    On Error GoTo EH
    
    ' 設定がオフ(False)の場合は何もしない（通常のクリック開封を許可）
    If Not modMailOpen.IsFeatureEnabled() Then Exit Sub
    
    ' 新規作成や返信・転送など、まだ送信されていないアイテム（下書き含む）はポップアップを許可
    If myItem.Sent = False Then Exit Sub
    
    ' ショートカット（独自マクロ）経由で開こうとしている場合は許可
    If modMailOpen.g_AllowOpen = True Then Exit Sub
    
    ' 上記以外（通常のマウスクリックでの開封）は完全に無反応にする（キャンセル）
    Cancel = True
    Exit Sub

EH:
    ' エラー時は安全のために処理を抜ける
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
