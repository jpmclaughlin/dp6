VERSION 5.00
Begin VB.Form frmCOMERROR 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Suggerimento del giorno"
   ClientHeight    =   4305
   ClientLeft      =   2310
   ClientTop       =   2055
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4830
      Top             =   630
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   585
      Left            =   1770
      TabIndex        =   0
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check siemens PG/PC interface settings, com. cable and plc."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   810
      TabIndex        =   2
      Top             =   2160
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   3120
      X2              =   3360
      Y1              =   1350
      Y2              =   1410
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   2460
      X2              =   2700
      Y1              =   1470
      Y2              =   1530
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   3120
      X2              =   3360
      Y1              =   1530
      Y2              =   1470
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   2460
      X2              =   2700
      Y1              =   1410
      Y2              =   1350
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   2430
      X2              =   3390
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   1
      Left            =   1650
      Picture         =   "frmCOMERROR.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   0
      Left            =   3510
      Picture         =   "frmCOMERROR.frx":0ABA
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   210
      Picture         =   "frmCOMERROR.frx":1574
      Stretch         =   -1  'True
      Top             =   690
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COM ERROR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1770
      TabIndex        =   1
      Top             =   450
      Width           =   3915
   End
End
Attribute VB_Name = "frmCOMERROR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Resize()
Move 5000, 4000
End Sub

Private Sub Timer1_Timer()
Static one

one = Not one

If one Then
  Me.BackColor = &HFFFF&
Else
   Me.BackColor = &HFF&
End If
End Sub
