VERSION 5.00
Begin VB.Form DialogChangeOrder 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4320
   ClientLeft      =   2760
   ClientTop       =   3465
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextInt 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   510
      TabIndex        =   4
      Text            =   "12345678901"
      Top             =   1590
      Width           =   6825
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
      Left            =   5490
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   2115
   End
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   2115
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vuoi annullare il cambio ordine"
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
      Height          =   1035
      Left            =   300
      TabIndex        =   5
      Top             =   1410
      Width           =   7335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vuoi introdurre in macchina l'ordine"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      TabIndex        =   3
      Top             =   1200
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Richiesta cambio ordine"
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
      Height          =   1035
      Left            =   300
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "DialogChangeOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Risposta As Boolean
Public Prenotato As Boolean
Private Sub CancelButton_Click()
   Risposta = False
   Me.Hide
End Sub

Private Sub Form_Activate()
   Label3.Visible = Prenotato
   Label1.Visible = Not Prenotato
   Label2.Visible = Not Prenotato
   TextInt.Visible = Not Prenotato
   Label1 = Param.Text("000000075")
   Label2 = Param.Text("000000076")
   Label3 = Param.Text("000000082")
   CancelButton.Caption = Param.Text("cancella")
   OKButton.Caption = Param.Text("ok")
End Sub

Private Sub Form_Load()
   Label1 = Param.Text("000000075")
   Label2 = Param.Text("000000076")
   Label3 = Param.Text("000000082")
   CancelButton.Caption = Param.Text("cancella")
   OKButton.Caption = Param.Text("ok")
End Sub

Private Sub OKButton_Click()
   Risposta = True
   Me.Hide
End Sub
