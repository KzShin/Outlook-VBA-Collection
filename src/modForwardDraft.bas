Attribute VB_Name = "modForwardDraft"
Option Explicit

' ==============================================================================
' Module: modForwardDraft
' Description: 選択したメールを指定の宛先へ転送する下書きを作成する（テキスト形式強制）
' Dependencies: Scripting.FileSystemObject, ADODB.Stream, modLogger, frmSelectDest
' Configuration: %APPDATA%\OutlookVBA\config.ini ([ForwardMail] Section)
' ==============================================================================

Private g_RunId As String

Private Sub SetRunId(ByVal id As String)
    g_RunId = Trim$(id)
End Sub

Private Sub Log(ByVal msg As String)
    modLogger.Log "ForwardDraft", msg
End Sub

' ==============================================================================
' [Public] 公開インターフェース
' ==============================================================================

Public Sub CreateForwardDraft()
    On Error GoTo EH
    
    Dim rid As String
    rid = Format(Now, "yymmdd-hhnnss") & "-FWD"
    SetRunId rid
    modLogger.SetRunId rid
    
    Log "=== START 転送メール作成処理 ==="
    
    ' 1. 対象アイテムの確認
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "転送するメールを選択してください。", vbExclamation
        Log "選択アイテムなし"
        GoTo FIN
    End If
    
    Dim objItem As Object
    Set objItem = Application.ActiveExplorer.Selection(1)
    
    If objItem.Class <> olMail Then
        MsgBox "選択されたアイテムはメールではありません。", vbExclamation
        Log "対象外アイテム: " & TypeName(objItem)
        GoTo FIN
    End If
    
    Dim mailItem As Outlook.MailItem
    Set mailItem = objItem
    Log "対象メール: " & mailItem.Subject
    
    ' 2. 転送先の読み込み
    Dim dictDests As Object
    Set dictDests = LoadForwardDestinations()
    
    If dictDests.Count = 0 Then
        MsgBox "config.ini に転送先 ([ForwardMail] セクション) が設定されていません。", vbExclamation
        Log "転送先設定なし"
        GoTo FIN
    End If
    
    ' 3. 転送先の決定
    Dim targetName As String
    Dim targetEmail As String
    
    If dictDests.Count = 1 Then
        ' 1件しかない場合は自動的にそれを選択
        Dim keys As Variant
        keys = dictDests.keys
        targetName = keys(0)
        targetEmail = dictDests(keys(0))
        Log "単一設定のため自動選択: " & targetName
    Else
        ' 複数ある場合はユーザーフォーム(frmSelectDest)を表示
        Log "複数設定あり、選択画面を表示"
        
        Dim frm As frmSelectDest
        Set frm = New frmSelectDest
        
        Dim key As Variant
        For Each key In dictDests.keys
            frm.lstDest.AddItem key
            frm.lstDest.List(frm.lstDest.ListCount - 1, 1) = dictDests(key)
        Next key
        
        frm.Show ' モーダルで表示
        
        If frm.IsCancelled Then
            Log "ユーザーによってキャンセルされました"
            Unload frm
            GoTo FIN
        End If
        
        targetName = frm.SelectedName
        targetEmail = frm.SelectedEmail
        Unload frm
        
        Log "ユーザー選択: " & targetName & " <" & targetEmail & ">"
    End If
    
    ' 4. 転送メールの生成と下書き保存
    CreateAndSaveDraft mailItem, targetName, targetEmail
    
FIN:
    Log "=== END 転送メール作成処理 ==="
    modLogger.SetRunId "NoID"
    Exit Sub
EH:
    Log "ERROR #" & Err.Number & " : " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    modLogger.SetRunId "NoID"
End Sub

' ==============================================================================
' [Logic] 下書き作成
' ==============================================================================

Private Sub CreateAndSaveDraft(ByVal origMail As Outlook.MailItem, ByVal destName As String, ByVal destEmail As String)
    On Error GoTo EH
    
    Dim fwMail As Outlook.MailItem
    Set fwMail = origMail.Forward
    
    ' 宛先の設定
    fwMail.To = destEmail
    
    ' HTMLメール等であってもテキスト形式に強制変換 (olFormatPlain = 1)
    fwMail.BodyFormat = olFormatPlain
    
    ' 本文の先頭に定型文を挿入（「様」は付けない）
    fwMail.Body = destName & vbCrLf & vbCrLf & _
                  "転送します。" & vbCrLf & vbCrLf & _
                  fwMail.Body
    
    ' 下書きとして保存
    fwMail.Save
    
    Log "下書きへ保存完了: " & fwMail.Subject & " (テキスト形式変換済)"
    MsgBox "転送メールを下書きに保存しました。" & vbCrLf & "宛先: " & destName, vbInformation, "完了"
    
    Exit Sub
EH:
    Err.Raise Err.Number, "CreateAndSaveDraft", Err.Description
End Sub

' ==============================================================================
' [Config] 設定読み込み
' ==============================================================================

Private Function LoadForwardDestinations() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim configPath As String
    configPath = Environ$("APPDATA") & "\OutlookVBA\config.ini"
    
    If Not fso.FileExists(configPath) Then
        Set LoadForwardDestinations = dict
        Exit Function
    End If
    
    ' UTF-8 読み込み
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile configPath
    
    Dim allText As String: allText = stm.ReadText(-1)
    stm.Close
    
    Dim lines() As String: lines = Split(Replace(allText, vbCrLf, vbLf), vbLf)
    Dim i As Long, lineText As String
    Dim currentSection As String
    Dim parts() As String, destParts() As String
    
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
            
            ' セクション判定
            If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
                currentSection = LCase$(Mid$(lineText, 2, Len(lineText) - 2))
            
            ' [ForwardMail] セクションのみ処理
            ElseIf currentSection = "forwardmail" Then
                parts = Split(lineText, "=", 2)
                If UBound(parts) = 1 Then
                    ' 値をカンマで分割 (表示名 , メールアドレス)
                    destParts = Split(parts(1), ",")
                    If UBound(destParts) >= 1 Then
                        ' Dictionary に追加 (Key:表示名, Value:アドレス)
                        dict.Add Trim$(destParts(0)), Trim$(destParts(1))
                    End If
                End If
            End If
            
        End If
    Next i
    
    Set LoadForwardDestinations = dict
End Function
