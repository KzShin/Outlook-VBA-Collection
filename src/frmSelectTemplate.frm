VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectTemplate 
   Caption         =   "テンプレートの選択"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6560
   OleObjectBlob   =   "frmSelectTemplate.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 呼び出し元（標準モジュール）に返すための公開変数
Public SelectedFilePath As String
Public IsCancelled As Boolean

Private Sub UserForm_Initialize()
    IsCancelled = False
    ' リストボックスの設定（隠し列にフルパス、表示列にファイル名を持たせる）
    With lstTemplates
        .ColumnCount = 2
        .ColumnWidths = "200 pt; 0 pt" ' 2列目（パス）は見えないように幅0にする
    End With
End Sub

Private Sub btnOK_Click()
    If lstTemplates.ListIndex >= 0 Then
        ' 選択されたアイテムの2列目（フルパス）を取得
        SelectedFilePath = lstTemplates.List(lstTemplates.ListIndex, 1)
        Me.Hide
    Else
        MsgBox "テンプレートを選択してください。", vbExclamation, "確認"
    End If
End Sub

Private Sub btnCancel_Click()
    IsCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        IsCancelled = True
        Me.Hide
    End If
End Sub

