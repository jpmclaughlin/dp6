VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form StopLogForm 
   Caption         =   "StopLog"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   13395
   WindowState     =   2  'Maximized
   Begin VB.Data DailyLogData 
      Caption         =   "DailyLogData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\log.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DailyStop"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data StopLogData 
      Caption         =   "StopLogData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\log.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "StopLog"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid StopLogDBGrid 
      Bindings        =   "StopLogForm.frx":0000
      Height          =   3855
      Left            =   0
      OleObjectBlob   =   "StopLogForm.frx":001A
      TabIndex        =   0
      Top             =   0
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DailyLogDBGrid 
      Bindings        =   "StopLogForm.frx":1243
      Height          =   3015
      Left            =   0
      OleObjectBlob   =   "StopLogForm.frx":125E
      TabIndex        =   1
      Top             =   4320
      Width           =   11655
   End
End
Attribute VB_Name = "StopLogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ricordarsi di chiudere e riaprire gli oggetti *Data
' qui utilizzati nella funzione "CompactLogDatabase"
' nel modulo "OrdersForm"


'****************************************************
' Funzione per impedire la chiusura del form da parte
' dell'utente
'****************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormControlMenu Then
        On Error Resume Next
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

