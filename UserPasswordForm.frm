VERSION 5.00
Begin VB.Form UserPasswordForm 
   Caption         =   "User password"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox PasswordText 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2160
      TabIndex        =   2
      Top             =   780
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   375
      TabIndex        =   1
      Top             =   780
      Width           =   1140
   End
   Begin VB.Label PasswordLabel 
      Caption         =   "Password :"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "UserPasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If PasswordText.Text = Param.Text("UserPassword") Then
        PasswordText.Text = ""
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox Param.Text("PasswordNotValid"), , ""
        PasswordText.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    LoginSucceeded = False
End Sub

Private Sub Form_Load()
    LoginSucceeded = False
End Sub

