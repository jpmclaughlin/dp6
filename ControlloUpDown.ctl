VERSION 5.00
Begin VB.UserControl ControlloUpDown 
   BackColor       =   &H0080FFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   FillColor       =   &H00FFFFFF&
   MaskColor       =   &H00FFFF80&
   ScaleHeight     =   930
   ScaleWidth      =   3255
   ToolboxBitmap   =   "ControlloUpDown.ctx":0000
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
      Height          =   900
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   1000
   End
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
      Height          =   900
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   890
      Left            =   1020
      TabIndex        =   2
      Top             =   30
      Width           =   1215
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   435
         Top             =   270
      End
      Begin VB.Label ValueDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "99.9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   75
         TabIndex        =   3
         Top             =   120
         Width           =   1080
      End
   End
End
Attribute VB_Name = "ControlloUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public value As Double
Public Step As Double
Public LimMax As Double
Public LimMin As Double
Public Decimali As Integer
Private Aumenta As Boolean
Public Occupato As Boolean
Public Cliccato As Boolean

Public Sub Refresh()
    If Step = 10 Then value = Round(value / 10) * 10
    If value < LimMin Then value = LimMin
    If value > LimMax Then value = LimMax
    If Decimali = 0 Then
        ValueDisplay.Caption = Format(value, "0")
    Else
        If Decimali = 1 Then
            ValueDisplay.Caption = Format(value, "0.0")
        Else
            ValueDisplay.Caption = Format(value, "0.00")
        End If
    End If
End Sub

Private Sub ComAumenta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Occupato = True
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Aumenta = True
        CicloAumenta
End Sub

Private Sub ComDiminuisce_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Occupato = True
        Timer1.Interval = 1000
        Timer1.Enabled = True
        Aumenta = False
        CicloDiminuisce
End Sub

Private Sub ComAumenta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Occupato = True
    Timer1.Enabled = False
    Cliccato = True
    Parent.AggiornaDaUpDownControl
    Cliccato = False
    Occupato = False
End Sub

Private Sub ComDiminuisce_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Occupato = True
    Timer1.Enabled = False
    Cliccato = True
    Parent.AggiornaDaUpDownControl
    Cliccato = False
    Occupato = False
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
    If value < LimMax Then
        value = value + Step
        If value > LimMax Then
            value = LimMax
        End If
        Occupato = True
        Refresh
    End If
End Sub

Private Sub CicloDiminuisce()
    Timer1.Interval = 200
    If value > LimMin Then
        value = value - Step
        If value < LimMin Then
            value = LimMin
        End If
        Occupato = True
        Refresh
    End If
End Sub


