VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form AlarmTextForm 
   Caption         =   "Alarm text"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CloseCommand 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12270
      TabIndex        =   4
      Top             =   8670
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid AlarmTextDBGrid 
      Align           =   1  'Align Top
      Bindings        =   "AlarmTextForm.frx":0000
      Height          =   8460
      Left            =   0
      OleObjectBlob   =   "AlarmTextForm.frx":001C
      TabIndex        =   0
      Top             =   0
      Width           =   6825
   End
   Begin VB.Data AlarmTextData 
      Caption         =   "AlarmText"
      Connect         =   "Access"
      DatabaseName    =   "..\target\alarms.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4710
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PlcText"
      Top             =   9030
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Type   1 as TRUE"
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   9150
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Type   0 as FALSE"
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   8910
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Type (Ctrl+Enter) to insert multi-line alarm text"
      Height          =   255
      Left            =   210
      TabIndex        =   1
      Top             =   8670
      Width           =   3615
   End
End
Attribute VB_Name = "AlarmTextForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseCommand_Click()
    Unload Me
End Sub

' è necessario fermare l'aggiornamento della finestra allarmi
' mentre si modificano gli allarmi per evitare conflitti
' nell'accesso al database
Private Sub Form_Load()
'    AlarmTextData.DataBaseName = AlarmForm.AlarmDataBaseName
    frmKernel.Timer1.Enabled = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmKernel.Timer1.Enabled = True
End Sub
