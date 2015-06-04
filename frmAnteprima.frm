VERSION 5.00
Begin VB.Form frmAnteprima 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmAnteprima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    tOp = 2000
    left = 2000

End Sub


Private Sub Form_Resize()
    DrawAnteprima
End Sub
