VERSION 5.00
Begin VB.UserControl ControlloSelettore 
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   ScaleHeight     =   1860
   ScaleWidth      =   1635
   ToolboxBitmap   =   "ControlloSelettore.ctx":0000
   Begin VB.Label Label1 
      Caption         =   "Selettore"
      Height          =   210
      Left            =   465
      TabIndex        =   0
      Top             =   90
      Width           =   675
   End
End
Attribute VB_Name = "ControlloSelettore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public db As DBClass
Public WordNum As Integer
Public BitMask As Integer

Private Sub UserControl_Click()
    db.Bit(WordNum, BitMask) = Not db.Bit(WordNum, BitMask)
    
    If db.Bit(WordNum, BitMask) Then
        Label1.Caption = "on"
    Else
        Label1.Caption = "off"
    End If
End Sub

