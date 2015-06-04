VERSION 5.00
Begin VB.UserControl Barra2 
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ScaleHeight     =   1125
   ScaleWidth      =   15360
   Begin VB.Frame Frame2 
      BackColor       =   &H009A9A9A&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15465
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   210
         Top             =   240
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Orders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   0
         Left            =   150
         Picture         =   "Barra2.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Layout"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   2
         Left            =   1410
         Picture         =   "Barra2.ctx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   11
         Left            =   2670
         Picture         =   "Barra2.ctx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "WB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   9
         Left            =   3930
         Picture         =   "Barra2.ctx":0755
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Threads"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   8
         Left            =   5190
         Picture         =   "Barra2.ctx":08C4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Bundle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   3
         Left            =   6450
         Picture         =   "Barra2.ctx":0990
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Alarms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   6
         Left            =   14010
         Picture         =   "Barra2.ctx":0AC0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Storage"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   5
         Left            =   8970
         Picture         =   "Barra2.ctx":0DCA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "chamf. m/c"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   7
         Left            =   10230
         Picture         =   "Barra2.ctx":0ED6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Taglio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   10
         Left            =   11490
         Picture         =   "Barra2.ctx":0FAA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Washing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   13
         Left            =   12750
         Picture         =   "Barra2.ctx":14A6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton Pulsante 
         BackColor       =   &H80000000&
         Caption         =   "Straps"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   4
         Left            =   7710
         Picture         =   "Barra2.ctx":15BC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1245
      End
   End
End
Attribute VB_Name = "Barra2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum IndicePuls
     pOrdini = 1
     pstorico = 2
     pMappa = 3
     pPacco = 4
     pRegge = 5
     pPesa = 6
     pAllarmi = 7
     pEntrata = 8
     pSmussatrice = 9
     pFilettatura = 10
     pWalkingBeam = 11
     pTaglio = 12
     pLavaggio = 14
End Enum

Const PulsanteID_Max = 14
Const PulsanteID_Alarm = 6

Private m_Selezionato As Integer
Private m_Allarme As Boolean

Public Event RipetizioneTasto()
Public Event PulsantePremuto(ByVal Index As IndicePuls)

Private Sub Pulsante_Click(Index As Integer)
   Dim i As Integer
   
   On Error Resume Next
   If m_Selezionato <> Index Then
      RaiseEvent PulsantePremuto(Index + 1)
      For i = 0 To PulsanteID_Max
         Pulsante(i).BackColor = &H80000000
      Next
      Pulsante(Index).BackColor = vbGreen
   Else
      RaiseEvent RipetizioneTasto
   End If
End Sub

Property Get Selezionato() As Integer
   Selezionato = m_Selezionato
End Property

Property Let Selezionato(ByVal Index As Integer)
   Dim i As Integer
   
   On Error Resume Next
   m_Selezionato = Index - 1
   For i = 0 To PulsanteID_Max
       Pulsante(i).BackColor = &H80000000
   Next
   Pulsante(m_Selezionato).BackColor = vbGreen
End Property

Private Sub Timer1_Timer()
   Static Lamp As Boolean
   
   Lamp = Not Lamp
   If Lamp Then
      Pulsante(PulsanteID_Alarm).BackColor = vbRed
   Else
      Pulsante(PulsanteID_Alarm).BackColor = &H80000000
   End If
End Sub

Private Sub UserControl_Initialize()
   Selezionato = 1
   Timer1.Enabled = False
   Timer1.Interval = 500
End Sub

Property Let Allarme(ByVal value As Boolean)
    m_Allarme = value
    Timer1.Interval = 500
    Timer1.Enabled = value
    If value = False Then
      If m_Selezionato = 6 Then
         Pulsante(PulsanteID_Alarm).BackColor = vbGreen
      Else
         Pulsante(PulsanteID_Alarm).BackColor = &H80000000
      End If
    End If
End Property
Property Get Allarme() As Boolean
   Allarme = m_Allarme
End Property

Sub Refresh_lingua()
   Pulsante(0).Caption = Param.Text("000000055")
   Pulsante(11).Caption = Param.Text("entrata")
   Pulsante(3).Caption = Param.Text("000000056")
   Pulsante(4).Caption = Param.Text("strap")
   Pulsante(5).Caption = Param.Text("stoccaggio")
   Pulsante(13).Caption = Param.Text("lavaggiopag")
   Pulsante(6).Caption = Param.Text("allarmi")
End Sub
