VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelp 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   10515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser webPreview 
      Height          =   9555
      Left            =   60
      TabIndex        =   3
      Top             =   900
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   16854
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   7020
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblContesto 
      BackStyle       =   0  'Transparent
      Caption         =   "lblContesto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2160
      TabIndex        =   1
      Top             =   300
      Width           =   4995
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Contesto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   2415
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   60
      TabIndex        =   4
      Top             =   90
      Width           =   7725
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' DB locale di test
Private DBTest As DBClass
' variabile
Private m_numPagine As Integer
Private m_PaginaCorrente As Integer
Private m_Percorso As String
Public NomeFile As String
Public Errori As Boolean

Property Let Contesto(ByVal Testo As String)
   On Error Resume Next
   lblContesto = Testo
End Property

Private Sub Form_Deactivate()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim PathFile As String
   Dim a() As String
   Dim i As Integer
    
    Label1.caption = Param.Text("Contesto")
    If Not Errori Then
       PathFile = LogComPath
    Else
       PathFile = HelpPath & NomeFile
    End If
    webPreview.Navigate PathFile
End Sub
Private Sub Label2_Click()
    Hide
    Unload Me
End Sub

