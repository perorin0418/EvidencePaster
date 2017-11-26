VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EP_RefEdit 
   Caption         =   "セルを選択"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "EP_RefEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "EP_RefEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    EP_module.CellRange = RefEdit.text
End Sub
