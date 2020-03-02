VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cookie Clicker Clicker"
   ClientHeight    =   855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2895
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnQuit_Click()
    bQuit = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    bQuit = True
End Sub
