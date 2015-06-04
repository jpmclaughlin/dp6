VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form OrdersLogForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Storico ordini"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11880
   WindowState     =   1  'Minimized
   Begin VB.CommandButton CommandPgDown 
      Caption         =   "PgDown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   13815
      MaskColor       =   &H0000FFFF&
      TabIndex        =   5
      Top             =   6390
      UseMaskColor    =   -1  'True
      Width           =   1380
   End
   Begin VB.CommandButton CommandPgUp 
      Caption         =   "PgUp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   13815
      MaskColor       =   &H0000FFFF&
      TabIndex        =   4
      Top             =   15
      UseMaskColor    =   -1  'True
      Width           =   1380
   End
   Begin VB.CommandButton ResetCommand 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   8805
      Width           =   3615
   End
   Begin VB.CommandButton DailyLogCommand 
      Caption         =   "Daily log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   9900
      MaskColor       =   &H0000FFFF&
      TabIndex        =   2
      Top             =   8865
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton BundlesLogCommand 
      Caption         =   "Bundles log"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   7470
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   8865
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Data OrdersLogData 
      Caption         =   "OrdersLogData"
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
      RecordSource    =   "OrdersLog"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid OrdersLogDBGrid 
      Bindings        =   "OrdersLogForm.frx":0000
      Height          =   8550
      Left            =   60
      OleObjectBlob   =   "OrdersLogForm.frx":001C
      TabIndex        =   0
      Top             =   45
      Width           =   13635
   End
End
Attribute VB_Name = "OrdersLogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'' ricordarsi di chiudere e riaprire gli oggetti *Data
'' qui utilizzati nella funzione "CompactLogDatabase"
'' nel modulo "OrdersForm"
'
'Private Sub CommandPgUp_Click()
'    OrdersLogDBGrid.Scroll 0, 20
'End Sub
'
'Private Sub CommandPgDown_Click()
'    OrdersLogDBGrid.Scroll 0, -20
'    OrdersLogDBGrid.Scroll 0, -1
'End Sub
'
'Private Sub Form_Activate()
'    OrdersLogData.Refresh
'End Sub
'
'Private Sub Form_Load()
'    Me.Caption = Param.Text("OrdersLogCommand")
'  '  OrdersLogDBGrid.Columns(0).Caption = PrinterForm.PrintField(DisplayFirstFieldPosition).Caption
'  '  OrdersLogDBGrid.Columns(1).Caption = Param.Text("Date")
'  '  OrdersLogDBGrid.Columns(2).Caption = Param.Text("BundleNum")
'  '  OrdersLogDBGrid.Columns(3).Caption = Param.Text("BundleWeight") & Unit.KgString
'  '  OrdersLogDBGrid.Columns(4).Caption = Param.Text("Tubes")
'  '  OrdersLogDBGrid.Columns(5).Caption = Param.Text("BundleLength") & Unit.mString
'
'    ResetCommand.Caption = Param.Text("LogReset")
'    BundlesLogCommand.Caption = Param.Text("BundlesLogCommand")
'    DailyLogCommand.Caption = Param.Text("DailyLogCommand")
'
'End Sub
'
''****************************************************
'' Funzione per impedire la chiusura del form da parte
'' dell'utente
''****************************************************
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If UnloadMode <> vbFormControlMenu Then
'        On Error Resume Next
'        Cancel = False
'    Else
'        Cancel = True
'    End If
'End Sub
'
'' cancellazione storico
'Private Sub ResetCommand_Click()
'   BundlesLogForm.ResetCommand_Click
'End Sub
'
'Private Sub BundlesLogCommand_Click()
'    On Error Resume Next
'    BundlesLogForm.Show
'    BundlesLogForm.ZOrder (0)
'    BundlesLogForm.BundlesLogData.Refresh
'End Sub
'
'Private Sub DailyLogCommand_Click()
'    On Error Resume Next
'    DailyLogForm.Show
'    DailyLogForm.ZOrder (0)
'    DailyLogForm.DailyData.Refresh
'End Sub
