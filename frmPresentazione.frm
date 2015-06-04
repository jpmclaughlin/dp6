VERSION 5.00
Begin VB.Form frmPresentazione 
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin dp6.vbalProgressBar PBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   10440
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   661
      Picture         =   "frmPresentazione.frx":0000
      BackColor       =   12582912
      ForeColor       =   0
      Appearance      =   0
      BarPicture      =   "frmPresentazione.frx":001C
      BarPictureMode  =   0
      BackPictureMode =   0
      Value           =   50
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM INITIALIZE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   645
      Left            =   150
      TabIndex        =   1
      Top             =   10020
      Width           =   5325
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM INITIALIZE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   9990
      Width           =   5325
   End
End
Attribute VB_Name = "frmPresentazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Operazione As Integer

Private Sub Form_Activate()
    If App.PrevInstance Then
       MsgBox " THIS PROGRAM IS ALREADY RUNNING...", vbCritical, "DP6 - ERROR"
       End
    Else
       Load frmKernel
    End If
End Sub

Private Sub Form_Load()
   Picture = LoadPicture("..\bitmap\presentazionemair.jpg")
   Show
   WindowState = vbMaximized
   ZOrder
End Sub

