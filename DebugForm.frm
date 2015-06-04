VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form DebugForm 
   Caption         =   "Debug"
   ClientHeight    =   5025
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Data ExitData 
      Caption         =   "ExitData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\plc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   9840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DB403"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data ActualBundleData 
      Caption         =   "ActualBundleData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\plc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DB411"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data FutureBundleData 
      Caption         =   "FutureBundleData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\plc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DB401"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data ActualStrapData 
      Caption         =   "ActualStrapData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\plc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DB417"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data FutureStrapData 
      Caption         =   "FutureStrapData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\plc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DB402"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data AlarmData 
      Caption         =   "AlarmData"
      Connect         =   "Access"
      DatabaseName    =   "..\target\plc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DB400"
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid ActualStrapDBGrid 
      Bindings        =   "DebugForm.frx":0000
      Height          =   8055
      Left            =   5880
      OleObjectBlob   =   "DebugForm.frx":001E
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid FutureStrapDBGrid 
      Bindings        =   "DebugForm.frx":0A1B
      Height          =   8055
      Left            =   4080
      OleObjectBlob   =   "DebugForm.frx":0A39
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DebugForm.frx":1436
      Height          =   3375
      Left            =   9840
      OleObjectBlob   =   "DebugForm.frx":144E
      TabIndex        =   4
      Top             =   4320
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid ExitDBGrid 
      Bindings        =   "DebugForm.frx":1E3D
      Height          =   3735
      Left            =   9840
      OleObjectBlob   =   "DebugForm.frx":1E54
      TabIndex        =   5
      Top             =   0
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid FutureBundleDBGrid 
      Bindings        =   "DebugForm.frx":2842
      Height          =   8055
      Left            =   0
      OleObjectBlob   =   "DebugForm.frx":2861
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid ActualBundleDBGrid 
      Bindings        =   "DebugForm.frx":325F
      Height          =   8055
      Left            =   1800
      OleObjectBlob   =   "DebugForm.frx":327E
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu Refresh 
      Caption         =   "Refresh"
   End
End
Attribute VB_Name = "DebugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Refresh_Click()
    AlarmData.Refresh
    'FutureEntryData.Refresh
    'ActualEntryData.Refresh
    'FutureFlushData.Refresh
    'ActualFlushData.Refresh
    'FutureFacerData.Refresh
    'ActualFacerData.Refresh
    FutureStrapData.Refresh
    ActualStrapData.Refresh
    FutureBundleData.Refresh
    ActualBundleData.Refresh
    ExitData.Refresh
End Sub

