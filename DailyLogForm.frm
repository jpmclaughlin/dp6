VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DailyLogForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Produzione giornaliera"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   12015
   WindowState     =   1  'Minimized
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
      Width           =   3375
   End
   Begin VB.CommandButton OrdersLogCommand 
      Caption         =   "Orders log"
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
      Left            =   9960
      MaskColor       =   &H0000FFFF&
      TabIndex        =   1
      Top             =   8865
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Data DailyData 
      Caption         =   "DailyData"
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
      RecordSource    =   "DailyProd"
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DailyDBGrid 
      Bindings        =   "DailyLogForm.frx":0000
      Height          =   8550
      Left            =   60
      OleObjectBlob   =   "DailyLogForm.frx":0018
      TabIndex        =   0
      Top             =   45
      Width           =   13635
   End
End
Attribute VB_Name = "DailyLogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ricordarsi di chiudere e riaprire gli oggetti *Data
' qui utilizzati nella funzione "CompactLogDatabase"
' nel modulo "OrdersForm"

'Private Sub Form_Load()
'    Me.Caption = Param.Text("DailyLog")
'    DailyDBGrid.Columns(0).Caption = Param.Text("Date")
'    DailyDBGrid.Columns(1).Caption = Param.Text("BundleWeight") & Unit.KgString
'    DailyDBGrid.Columns(2).Caption = Param.Text("Tubes")
'    DailyDBGrid.Columns(3).Caption = Param.Text("BundleLength") & Unit.mString
'
'    ResetCommand.Caption = Param.Text("LogReset")
'    BundlesLogCommand.Caption = Param.Text("BundlesLogCommand")
'    OrdersLogCommand.Caption = Param.Text("OrdersLogCommand")
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
'Public Sub ResetCommand_Click()
'   If Not DailyData.Recordset.EOF And Not DailyData.Recordset.BOF Then
'        If MsgBox(Param.Text("AreYouSure"), vbYesNo, Param.Text("LogReset")) = vbYes Then
'            On Error Resume Next
'            DailyData.Recordset.MoveFirst
'            While Not DailyData.Recordset.EOF
'                DailyData.Recordset.Delete
'                DailyData.Recordset.MoveNext
'            Wend
'        End If
'    End If
'End Sub
'
'Private Sub BundlesLogCommand_Click()
'    On Error Resume Next
'    BundlesLogForm.Show
'    BundlesLogForm.ZOrder (0)
'    BundlesLogForm.BundlesLogData.Refresh
'End Sub
'
'Private Sub OrdersLogCommand_Click()
'    On Error Resume Next
'    OrdersLogForm.Show
'    OrdersLogForm.ZOrder (0)
'    OrdersLogForm.OrdersLogData.Refresh
'End Sub
