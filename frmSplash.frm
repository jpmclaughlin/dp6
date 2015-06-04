VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5160
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   510
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dp6 Starter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   6855
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prodotto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   2430
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6885
      TabIndex        =   4
      Top             =   2940
      Width           =   90
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Mair Research s.p.a"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4830
      Width           =   2415
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4620
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgLogo 
      Height          =   5145
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Tempo As Integer
Private Ritardo As Long

Option Explicit

Private Sub Form_Load()
    Dim RS_Param As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    
    On Error GoTo Errore
    
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Parameters.mdb;Persist Security Info=False"
    With RS_Param
         ' open the connection to the DB and refresh the data
         .Open "SELECT Valore FROM Parametri WHERE ID='Par206_RitardoPartenza'", cn, , adLockReadOnly, adCmdText
         If .EOF Then Ritardo = 5
         Ritardo = .Fields("Valore")
         .Close
    End With
    Set RS_Param = Nothing
    Set cn = Nothing
    
    On Error Resume Next
   Timer1.Enabled = False
   If Ritardo < 0 Or Ritardo > 32767 Then MsgBox "The value of starter delay is out of range", vbCritical, "DP 6.0 - Fatal Error"
    If Ritardo > 0 Then
       lblVersion.caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
       lblProductName.caption = "Datapack"
       Tempo = 0
       Label1 = Ritardo - Tempo
       Timer1.Enabled = True
   Else
        Load frmPresentazione
        Unload Me
   End If
   Exit Sub
   
Errore:
    Kill "..\target\production.mdb"
    FileCopy "..\target\productionCopy.mdb", "..\target\production.mdb"
    Open "..\target\LogErrori.txt" For Append As #2
    Print #2, Format(Now, "dd-mm-yyyy hh:mm:ss") & " Ripristinato production.mdb "
    Close #2
    End
End Sub

Private Sub Timer1_Timer()
    Tempo = Tempo + 1
    Label1 = Ritardo - Tempo
    If Tempo < Ritardo Then Exit Sub
    Load frmPresentazione
    Unload Me
End Sub

