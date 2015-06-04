VERSION 5.00
Begin VB.Form frmCOMERROR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "TemplateShape"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3090
      TabIndex        =   0
      Top             =   4020
      Width           =   1785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COM ERROR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2970
      TabIndex        =   2
      Top             =   450
      Width           =   3915
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   1410
      Picture         =   "frmComErrorROT.frx":0000
      Stretch         =   -1  'True
      Top             =   690
      Width           =   930
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   0
      Left            =   4710
      Picture         =   "frmComErrorROT.frx":0442
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   570
   End
   Begin VB.Image Image2 
      Height          =   570
      Index           =   1
      Left            =   2850
      Picture         =   "frmComErrorROT.frx":0EFC
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   242
      X2              =   306
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   244
      X2              =   260
      Y1              =   94
      Y2              =   90
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   2
      X1              =   288
      X2              =   304
      Y1              =   102
      Y2              =   98
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   3
      X1              =   244
      X2              =   260
      Y1              =   98
      Y2              =   102
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   4
      X1              =   288
      X2              =   304
      Y1              =   90
      Y2              =   94
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Check siemens PG/PC interface settings (CP_L2_1), com. cable and plc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   1890
      TabIndex        =   1
      Top             =   1950
      Width           =   4365
   End
End
Attribute VB_Name = "frmCOMERROR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1

Private Const PS_SOLID = 1

Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Sub RitagliaForm(ByVal Incremento As Single)

Dim rgn As Long
Dim old_rgn As Long
Dim X1 As Long, Y1 As Long
Dim X2 As Long, Y2 As Long
Dim X3 As Long, Y3 As Long
Dim Ret1 As Long, Ret2 As Long
Dim Ret As Long
Dim pen As Long
Dim old_pen As Long

Me.Cls

If Incremento > 1 Then Incremento = 1

X1 = 500 / 2 - (500 / 2 * Incremento)
Y1 = 300 / 2 - (300 / 2 * Incremento)
X2 = 500 / 2 + (550 / 2 * Incremento)
Y2 = 300 / 2 + (450 / 2 * Incremento)
X3 = 90 / 2 + (90 / 2 * Incremento)
Y3 = X3

Ret1 = CreateRoundRectRgn(X1, Y1, X2, Y2, X3, Y3)
Ret2 = SetWindowRgn(Me.hwnd, Ret1, True)
pen = CreatePen(PS_SOLID, 4, 0) ' largezza penna
old_pen = SelectObject(hdc, pen)
Ret = RoundRect(hdc, X1, Y1, X2, Y2, X3, Y3)
pen = SelectObject(hdc, old_pen)
Ret = DeleteObject(pen)
Me.Refresh

End Sub

Private Sub cmdOK_Click()
   On Error Resume Next
   Unload Me
End Sub

Private Sub Form_Load()
Dim Ret As Long

    'mantiene il form sempre in primo piano (alway on top)
    Ret = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)

End Sub
Private Sub Form_Resize()
Dim Start As Double
Dim temp As Single
Dim Delay As Single

'modificare, eventualmente, il valore dello Step e/o del Delay fino
'ad ottenere l'effetto di visualizzazione  voluta.

Delay = 0.01

For temp = 0.1 To 1.1 Step 0.1
   Call RitagliaForm(temp)
   Start = Timer
      Do While Timer < Start + 0.01
         DoEvents
      Loop
 Next temp

End Sub
Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'permette lo spostamento del form con il cursore del mouse

Dim ReturnVal As Long
If Button = 1 Then
     x = ReleaseCapture()
     ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
 End If
End Sub

