VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectDest 
   Caption         =   "転送先の選択"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6560
   OleObjectBlob   =   "frmSelectDest.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectDest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 呼び出し元（標準モジュール）に返すための公開変数
Public SelectedName As String
Public SelectedEmail As String
Public IsCancelled As Boolean

Private Sub IsDest_Click()

End Sub

' フォーム初期化時
Private Sub UserForm_Initialize()
    IsCancelled = False
    
    ' リストボックスを2列（表示名、アドレス）に設定
    With lstDest
        .ColumnCount = 2
        .ColumnWidths = "120 pt; 150 pt" ' 幅はお好みで調整してください
    End With
End Sub

' OKボタンクリック時
Private Sub btnOK_Click()
    If lstDest.ListIndex >= 0 Then
        SelectedName = lstDest.List(lstDest.ListIndex, 0)
        SelectedEmail = lstDest.List(lstDest.ListIndex, 1)
        Me.Hide ' フォームを非表示にして処理を標準モジュールへ戻す
    Else
        MsgBox "転送先を選択してください。", vbExclamation, "確認"
    End If
End Sub

' キャンセルボタンクリック時
Private Sub btnCancel_Click()
    IsCancelled = True
    Me.Hide
End Sub

' 「×」ボタンで閉じられたときの処理
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        IsCancelled = True
        Me.Hide
    End If
End Sub

