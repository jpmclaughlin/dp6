VERSION 5.00
Begin VB.Form TechPasswordForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service password "
   ClientHeight    =   5625
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   9915
   Icon            =   "TechPasswordForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3323.436
   ScaleMode       =   0  'User
   ScaleWidth      =   9309.648
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1693
      Left            =   810
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   3195
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1693
      Left            =   5835
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   3195
   End
   Begin VB.TextBox PasswordText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      IMEMode         =   3  'DISABLE
      Left            =   2827
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   4260
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2070
      Picture         =   "TechPasswordForm.frx":030A
      Top             =   1230
      Width           =   480
   End
   Begin VB.Label PasswordLabel 
      Alignment       =   2  'Center
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3157
      TabIndex        =   3
      Top             =   360
      Width           =   3645
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   1470
      Shape           =   3  'Circle
      Top             =   1050
      Width           =   1575
   End
End
Attribute VB_Name = "TechPasswordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public defPassWord As String

Private Sub cmdCancel_Click()
    PasswordText.Text = ""
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Trim(PasswordText.Text) = defPassWord Then
        PasswordText.Text = ""
        LoginSucceeded = True
        Me.Hide
    Else
        LoginSucceeded = False
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    PasswordText.SetFocus
    LoginSucceeded = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    LoginSucceeded = False
End Sub

Private Sub Form_Load()
    LoginSucceeded = False
End Sub

Private Sub PasswordText_Click()
    TOUCHKeyBoard.TextModifica.PasswordChar = "*"
    TOUCHKeyBoard.Dati = ""
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        PasswordText.Text = TOUCHKeyBoard.Dati
    End If
    TOUCHKeyBoard.TextModifica.PasswordChar = ""
End Sub
