VERSION 5.00
Begin VB.Form TOUCHNumericPad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10530
   ClientLeft      =   3015
   ClientTop       =   510
   ClientWidth     =   9240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   9240
   Begin VB.CommandButton KeyClear 
      BackColor       =   &H8000000B&
      Cancel          =   -1  'True
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8940
      Width           =   2340
   End
   Begin VB.CommandButton KeyEnter 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8940
      Width           =   2520
   End
   Begin VB.CommandButton KeyEsc 
      BackColor       =   &H000000FF&
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8940
      Width           =   2520
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "<< BS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Index           =   0
      Left            =   7020
      TabIndex        =   13
      Top             =   1260
      Width           =   1935
   End
   Begin VB.TextBox TextModifica 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   660
      TabIndex        =   12
      Top             =   1260
      Width           =   6195
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   12
      Left            =   5100
      TabIndex        =   11
      Top             =   7260
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   11
      Left            =   1980
      TabIndex        =   10
      Top             =   7260
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   3
      Left            =   5100
      TabIndex        =   9
      Top             =   2580
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   2
      Left            =   3540
      TabIndex        =   8
      Top             =   2580
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   1
      Left            =   1980
      TabIndex        =   7
      Top             =   2580
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   6
      Left            =   5100
      TabIndex        =   6
      Top             =   4140
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   5
      Left            =   3540
      TabIndex        =   5
      Top             =   4140
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   4
      Left            =   1980
      TabIndex        =   4
      Top             =   4140
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   9
      Left            =   5100
      TabIndex        =   3
      Top             =   5700
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   8
      Left            =   3540
      TabIndex        =   2
      Top             =   5700
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   7
      Left            =   1980
      TabIndex        =   1
      Top             =   5700
      Width           =   1500
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   10
      Left            =   3540
      TabIndex        =   0
      Top             =   7260
      Width           =   1500
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      Top             =   750
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   18
      Top             =   330
      Width           =   2145
   End
   Begin VB.Label LabelDati 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   660
      TabIndex        =   17
      Top             =   240
      Width           =   6195
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   9210
      X2              =   9210
      Y1              =   15
      Y2              =   10500
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000007&
      BorderWidth     =   4
      X1              =   -15
      X2              =   9225
      Y1              =   -15
      Y2              =   -15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   60
      X2              =   60
      Y1              =   0
      Y2              =   10500
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000007&
      BorderWidth     =   4
      X1              =   -30
      X2              =   9210
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      X1              =   9240
      X2              =   9240
      Y1              =   120
      Y2              =   10140
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      X1              =   9240
      X2              =   0
      Y1              =   10500
      Y2              =   10500
   End
End
Attribute VB_Name = "TOUCHNumericPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DatiConfermati As Boolean
Public Dati As Double
Public Decimali As Integer
Public ValoreMax As Variant
Public ValoreMin As Variant

Private Sub Form_Activate()
    DatiConfermati = False
    TextModifica.Text = ""
    Label1.Caption = "Min:" & ValoreMin
    Label2.Caption = "Max:" & ValoreMax
    If Decimali = 0 Then
        LabelDati.Caption = Format(Dati, "0")
    Else
        If Decimali = 1 Then
            LabelDati.Caption = Format(Dati, "0.0")
        Else
            LabelDati.Caption = Format(Dati, "0.00")
        End If
    End If
    TextModifica.SetFocus
    SendKeys "{END}"
End Sub

Private Sub KeyEnter_Click()
    ' Restituisce la stringa togliendo gli eventuali spazi iniziali e finali
    On Error Resume Next
        Dati = CDbl(TextModifica.Text)
        If Dati >= ValoreMin And Dati <= ValoreMax Then
           On Error GoTo 0
           DatiConfermati = True
           Me.Hide
        Else
           frmAvvisi.AvvisoBypass = False
           frmAvvisi.tOp = Me.tOp + Me.Height / 2
           frmAvvisi.left = Me.left + Me.Width / 2
           frmAvvisi.Show vbModal
        End If
End Sub

Private Sub KeyEsc_Click()
    DatiConfermati = False
    Me.Hide
End Sub

Private Sub KeyClear_Click()
    TextModifica.SetFocus
    TextModifica.Text = ""
    End Sub

Private Sub Key_Click(Index As Integer)
    TextModifica.SetFocus
    Select Case Index
        Case 0
            SendKeys "{BS}"
        Case 1
            SendKeys "1"
        Case 2
            SendKeys "2"
        Case 3
            SendKeys "3"
        Case 4
            SendKeys "4"
        Case 5
            SendKeys "5"
        Case 6
            SendKeys "6"
        Case 7
            SendKeys "7"
        Case 8
            SendKeys "8"
        Case 9
            SendKeys "9"
        Case 10
            SendKeys "0"
        Case 11
            SendKeys "{HOME}"
            If Mid(TextModifica.Text, 1, 1) = "-" Then
                SendKeys "{DEL}"
            Else
                SendKeys "-"
            End If
            SendKeys "{END}"
        Case 12
                SendKeys ","
    End Select
End Sub
