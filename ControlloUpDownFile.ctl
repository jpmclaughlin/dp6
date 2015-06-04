VERSION 5.00
Begin VB.UserControl ControlloUpDownFile 
   BackColor       =   &H0080FFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   FillColor       =   &H00FFFFFF&
   MaskColor       =   &H00FFFF80&
   ScaleHeight     =   3000
   ScaleWidth      =   1035
   ToolboxBitmap   =   "ControlloUpDownFile.ctx":0000
   Begin VB.CommandButton ComAumenta 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   1000
   End
   Begin VB.CommandButton ComDiminuisce 
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2025
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1000
      Left            =   15
      TabIndex        =   2
      Top             =   1005
      Width           =   975
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   300
         Top             =   330
      End
      Begin VB.Label ValueDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   45
         TabIndex        =   3
         Top             =   195
         Width           =   945
      End
   End
End
Attribute VB_Name = "ControlloUpDownFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Value As Double
Public Step As Double
Public LimMax As Double
Public LimMin As Double
Private Aumenta As Boolean

Public Sub Refresh()
        ValueDisplay = Value
End Sub

Private Sub ComAumenta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Aumenta = True
        CicloAumenta
End Sub

Private Sub ComDiminuisce_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Aumenta = False
        CicloDiminuisce
End Sub

Private Sub ComAumenta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    Parent.AggiornaDaUpDownControl
End Sub

Private Sub ComDiminuisce_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    Parent.AggiornaDaUpDownControl
End Sub

Private Sub Timer1_Timer()
    If Aumenta Then
        CicloAumenta
    Else
        CicloDiminuisce
    End If
End Sub

Private Sub CicloAumenta()
    Timer1.Interval = 200
    If Value < LimMax Then
        Value = Value + Step
        If Value > LimMax Then
            Value = LimMax
        End If
        ValueDisplay = Value
    End If
End Sub

Private Sub CicloDiminuisce()
    Timer1.Interval = 200
    If Value > LimMin Then
        Value = Value - Step
        If Value < LimMin Then
            Value = LimMin
        End If
        ValueDisplay = Value
    End If
End Sub


