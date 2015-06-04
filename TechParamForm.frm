VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form TechParamForm 
   Caption         =   "Technical parameters"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11370
   WindowState     =   2  'Maximized
   Begin VB.Data OptionsData 
      Caption         =   "OptionsData"
      Connect         =   "Access"
      DatabaseName    =   "C:\Usr\1961 Wirsbo\parameters.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "OptionsQuery"
      Top             =   5880
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Data NumbersData 
      Caption         =   "NumbersData"
      Connect         =   "Access"
      DatabaseName    =   "C:\Usr\1961 Wirsbo\parameters.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "NumbersQuery"
      Top             =   1440
      Visible         =   0   'False
      Width           =   5415
   End
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
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   7800
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid OptionsDBGrid 
      Bindings        =   "TechParamForm.frx":0000
      Height          =   2775
      Left            =   240
      OleObjectBlob   =   "TechParamForm.frx":001A
      TabIndex        =   1
      Top             =   4800
      Width           =   11295
   End
   Begin MSDBGrid.DBGrid NumbersDBGrid 
      Bindings        =   "TechParamForm.frx":0D47
      Height          =   4335
      Left            =   240
      OleObjectBlob   =   "TechParamForm.frx":0D61
      TabIndex        =   0
      Top             =   120
      Width           =   11295
   End
   Begin VB.Label Label2 
      Caption         =   "Type   0 as FALSE"
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Type   1 as TRUE"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   8040
      Width           =   1335
   End
End
Attribute VB_Name = "TechParamForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseCommand_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'    LanguageData.DataBaseName = Param.DataBaseName
'    LanguageData.RecordSource = Param.LanguageTableName
'    NumbersData.DataBaseName = Param.DataBaseName
'    OptionsData.DataBaseName = Param.DataBaseName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Download
End Sub

' Download dei parametri macchina
Public Sub Download()
'    ' dimensioni e posizioni in s
'    Dim DrainSensorWidth As Double
'    Dim StorageSensorWidth As Double
'    Dim SensorPosition As Double
'
'    On Error Resume Next
'        CommandDB.Word(CommandMap.Rallentamento) = Param.Number("StrapSlowDown") * 1000#
'        CommandDB.Word(CommandMap.OffsetReggiatura) = Param.Number("StrapOffset") * 1000#
'        CommandDB.Word(CommandMap.CoeffConteggioRegg) = Param.Number("StrapCoeff")
'        If Err.Number <> 0 Then MsgBox "Strap parameters error", vbOKOnly
'    On Error GoTo 0
'
'    On Error Resume Next
'        CommandDB.Word(CommandMap.CoeffConteggioCentr) = Param.Number("CoeffConteggioCentr")
'        DrainSensorWidth = Param.Number("DrainSensorWidth") / Param.Number("ChainSpeed")
'        StorageSensorWidth = Param.Number("StorageSensorWidth") / Param.Number("ChainSpeed")
'        SensorPosition = Param.Number("SensorPosition") / Param.Number("ChainSpeed")
'
'        #If S7Plc Then
'            CommandDB.Word(CommandMap.DrainSensorWidth) = DrainSensorWidth * 10#
'            CommandDB.Word(CommandMap.SensorPosition) = SensorPosition * 10#
'            CommandDB.Word(CommandMap.StorageSensorWidth) = StorageSensorWidth * 10#
'        #ElseIf ABPlc Then
'            CommandDB.Word(CommandMap.DrainSensorWidth) = DrainSensorWidth * 100#
'            CommandDB.Word(CommandMap.SensorPosition) = SensorPosition * 100#
'            CommandDB.Word(CommandMap.StorageSensorWidth) = StorageSensorWidth * 100#
'        #End If
'
'        If Err.Number <> 0 Then MsgBox "Storage parameters error", vbOKOnly
'    On Error GoTo 0
'
'    On Error Resume Next
'        CommandDB.Word(CommandMap.DistFascRegg) = Param.Number("DistFascRegg") * 1000#
'        CommandDB.Word(CommandMap.NumeroGiriAnelloTesta) = Param.Number("NumeroGiriAnelloTesta")
'        CommandDB.Word(CommandMap.NumeroGiriAnelloCoda) = Param.Number("NumeroGiriAnelloCoda")
'        CommandDB.Word(CommandMap.QuotaFascTesta) = Param.Number("QuotaFascTesta") * 1000#
'        CommandDB.Word(CommandMap.QuotaFascCoda) = Param.Number("QuotaFascCoda") * 1000#
'        If Err.Number <> 0 Then MsgBox "Bander parameters error", vbOKOnly
'    On Error GoTo 0
'



End Sub
