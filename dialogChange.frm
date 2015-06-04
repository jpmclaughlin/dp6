VERSION 5.00
Begin VB.Form dialogChange 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3150
      Width           =   2115
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000FF&
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5880
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3150
      Width           =   2115
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The data of the tube are changed"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   8445
   End
End
Attribute VB_Name = "dialogChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Risposta As Boolean
Public inChange As Boolean
Private Sub CancelButton_Click()
   Risposta = False
   inChange = False
   Me.Hide
End Sub

Private Sub Form_Activate()
   If inChange Then
      Label1 = Param.Text("000000077")
   Else
      Label1 = Param.Text("000000076")
   End If
   CancelButton.caption = Param.Text("cancella")
   OKButton.caption = Param.Text("ok")
End Sub

Private Sub Form_Load()
   If inChange Then
      Label1 = Param.Text("000000077")
   Else
      Label1 = Param.Text("000000076")
   End If
   CancelButton.caption = Param.Text("cancella")
   OKButton.caption = Param.Text("ok")
End Sub

Private Sub OKButton_Click()
   Risposta = True
   inChange = False
   Me.Hide
End Sub

