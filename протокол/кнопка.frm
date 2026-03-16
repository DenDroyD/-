VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClientType 
   Caption         =   "Выбрать тип Клиента"
   ClientHeight    =   1440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3780
   OleObjectBlob   =   "кнопка.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmClientType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Объявим публичную переменную для хранения результата
Public SelectedValue As String

Private Sub btnNew_Click()
    SelectedValue = "Новый"
    Me.Hide
End Sub

Private Sub btnRepeat_Click()
    SelectedValue = "Повторный"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Если закрыли форму крестиком - ничего не выбираем
    If SelectedValue = "" Then
        SelectedValue = "Не выбрано"
    End If
End Sub

