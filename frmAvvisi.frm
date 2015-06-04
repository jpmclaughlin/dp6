VERSION 5.00
Begin VB.Form frmAvvisi 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2130
      TabIndex        =   1
      Top             =   2520
      Width           =   1965
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning, must be used only with empty zone"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   300
      TabIndex        =   2
      Top             =   300
      Width           =   6225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valore non valido"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   6165
   End
End
Attribute VB_Name = "frmAvvisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

'Private Sub Command1_Click()
'   Hide
'   Unload Me
'End Sub
'
'Private Sub Form_Load()
'
'If AvvisoBypass Then
'   Label1.Visible = False
'   Label2.Visible = True
'Else
'   Label1.Visible = True
'   Label2.Visible = False
'End If
'
'   Label1.Caption = Param.Text("Valore")
'End Sub
'
'Private Sub Form_Resize()
'   Move 4500, 4000
'End Sub

Option Explicit
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Const HTCAPTION = 2
Const WM_NCLBUTTONDOWN = &HA1

Private Const PS_SOLID = 1

Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Public AvvisoBypass As Boolean
Sub RitagliaForm(ByVal Incremento As Single)

Dim rgn As Long
Dim old_rgn As Long
Dim X1 As Long, Y1 As Long
Dim X2 As Long, Y2 As Long
Dim X3 As Long, Y3 As Long
Dim Ret1 As Long, Ret2 As Long
Dim ret As Long
Dim pen As Long
Dim old_pen As Long

Me.Cls

If Incremento > 1 Then Incremento = 1

X1 = 400 / 2 - (400 / 2 * Incremento)
Y1 = 300 / 2 - (300 / 2 * Incremento)
X2 = 300 / 2 + (300 / 2 * Incremento)
Y2 = 300 / 2 + (300 / 2 * Incremento)
X3 = 90 / 2 + (90 / 2 * Incremento)
Y3 = X3

Ret1 = CreateRoundRectRgn(X1, Y1, X2, Y2, X3, Y3)
Ret2 = SetWindowRgn(Me.hwnd, Ret1, True)
pen = CreatePen(PS_SOLID, 4, 0) ' largezza penna
old_pen = SelectObject(hDC, pen)
ret = RoundRect(hDC, X1, Y1, X2, Y2, X3, Y3)
pen = SelectObject(hDC, old_pen)
ret = DeleteObject(pen)
Me.Refresh

End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim ret As Long

    'mantiene il form sempre in primo piano (alway on top)
    ret = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
    
End Sub
Private Sub Form_Resize()
Dim Start As Double
Dim temp As Single
Dim Delay As Single

'modificare, eventualmente, il valore dello Step e/o del Delay fino
'ad ottenere l'effetto di visualizzazione  voluta.

Delay = 0.01
Move 4500, 4000

If AvvisoBypass Then
   Label1.Visible = False
   Label2.Visible = True
Else
   Label1.Visible = True
   Label2.Visible = False
End If

   Label1.Caption = Param.Text("Valore")
   
For temp = 0.1 To 1.1 Step 0.1
   Call RitagliaForm(temp)
   Start = Timer
      Do While Timer < Start + 0.01
         DoEvents
      Loop
 Next temp


End Sub
Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'permette lo spostamento del form con il cursore del mouse

Dim ReturnVal As Long
If Button = 1 Then
     X = ReleaseCapture()
     ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
 End If
End Sub



