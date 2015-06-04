VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form AdjustForm 
   Caption         =   "Adjust "
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Mark axis counts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      TabIndex        =   9
      Top             =   0
      Width           =   2655
      Begin VB.Label MarkAxisCountDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Data MarkAxis 
      Caption         =   "MarkAxis"
      Connect         =   "Access"
      DatabaseName    =   "..\target\Adjust.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MarkAxis"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton SendMarkAxisCommand 
      Caption         =   "Send data"
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
      Left            =   2880
      TabIndex        =   6
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10680
      Top             =   7680
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
      Height          =   735
      Left            =   7200
      TabIndex        =   5
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Facer axis counts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2655
      Begin VB.Label FacerAxisCountDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton SendFacerAxisCommand 
      Caption         =   "Send data"
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
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Data FacerAxis 
      Caption         =   "FacerAxis"
      Connect         =   "Access"
      DatabaseName    =   "..\target\Adjust.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FacerAxis"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid FacerAxisDBGrid 
      Bindings        =   "AdjustForm.frx":0000
      Height          =   5895
      Left            =   0
      OleObjectBlob   =   "AdjustForm.frx":0014
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid MarkAxisDBGrid 
      Bindings        =   "AdjustForm.frx":09F6
      Height          =   5895
      Left            =   2880
      OleObjectBlob   =   "AdjustForm.frx":0A09
      TabIndex        =   7
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label MarkSendInProgressLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send in progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label FacerSendInProgressLabel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send in progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "AdjustForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
' Mappa dati del PLC
'*******************************************************
Private Const FacerAdjustDBNum As Long = 405  ' Numero del DB
Private Const MarkAdjustDBNum As Long = 406  ' Numero del DB

'********************
' Read and write area (non usata)
'********************
Private Const FacerAdjustReadWriteFirstWord As Long = 0
Private Const MarkAdjustReadWriteFirstWord As Long = 0

Private Const FacerAdjustReadWriteAmount As Long = 1
Private Const MarkAdjustReadWriteAmount As Long = 1

'***********
' Read area
'***********
Private Const FacerAdjustReadFirstWord As Long = 0 ' Numero della prima word
Private Const MarkAdjustReadFirstWord As Long = 0 ' Numero della prima word

Private Const FacerAdjustReadAmount As Long = 1 ' Numero di Word
Private Const MarkAdjustReadAmount As Long = 1 ' Numero di Word

Public Enum TAdjustReadMap
    AxisCount = 0
End Enum

'************
' Write area
'************
Private Const FacerAdjustWriteFirstWord As Long = 10
Private Const MarkAdjustWriteFirstWord As Long = 10
Private Const FacerAdjustWriteAmount As Long = 106
Private Const MarkAdjustWriteAmount As Long = 106
'dalla word 10 alla 115 si mettono i valori della tabella



Private Sub CloseCommand_Click()
    Unload Me
End Sub

'**************************************************
' La funzione di load del Form è utilizzata per il
' caricamento dei parametri e la inizializzazione
' della comunicazione con il PLC
'**************************************************
Private Sub Form_Load()
    
    ' inizializza la comunicazione con PLC
    ' NB: solo una istanza del plc può essere master (cioè il primo parametro = true)
    '     per convenzione è il plc degli allarmi
    FacerAdjustPlc.Initialize False, FacerAdjustDBNum, FacerAdjustReadFirstWord, FacerAdjustReadAmount, FacerAdjustWriteFirstWord, FacerAdjustWriteAmount, _
                       FacerAdjustReadWriteFirstWord, FacerAdjustReadWriteAmount
    FacerAdjustPlc.Enable = Param.Bit("PlcEnable")

    MarkAdjustPlc.Initialize False, MarkAdjustDBNum, MarkAdjustReadFirstWord, MarkAdjustReadAmount, MarkAdjustWriteFirstWord, MarkAdjustWriteAmount, _
                       MarkAdjustReadWriteFirstWord, MarkAdjustReadWriteAmount
    MarkAdjustPlc.Enable = Param.Bit("PlcEnable")


End Sub


'**************************************************
' Funzione di aggiornamento video e comunicazione
' con il plc da chiamare in background
'**************************************************
Private Sub Timer1_Timer()
    If FacerAdjustPlc.Dfa() Then
        ' confronto dei dati interni con quelli del PLC
        If DataChange() Then
            ' aggiornamento dei dati visualizzati
            DataUpload
        End If
    End If
    If MarkAdjustPlc.Dfa() Then
        ' confronto dei dati interni con quelli del PLC
        If DataChange() Then
            ' aggiornamento dei dati visualizzati
            DataUpload
        End If
    End If
    SignalUpdate
End Sub


'**************************************************
' Funzione per aggiornare le segnalazioni a video per l'operatore
'**************************************************
Private Sub SignalUpdate()
    FacerSendInProgressLabel.Visible = FacerAdjustPlc.OutWriteInProgress
    MarkSendInProgressLabel.Visible = MarkAdjustPlc.OutWriteInProgress
End Sub



'**********************************************************
'  funzioni per il trasferimento dati PLC<->Slider->Display
'**********************************************************


' facer count update
Private Sub FacerCountUpload()
    FacerAxisCountDisplay.Caption = FacerAdjustPlc.ReadWord(TAdjustReadMap.AxisCount)
End Sub
Private Function FacerCountChange() As Boolean
    If FacerAxisCountDisplay.Caption <> FacerAdjustPlc.ReadWord(TAdjustReadMap.AxisCount) Then
        FacerCountChange = True
    Else
        FacerCountChange = False
    End If
End Function


' mark count update
Private Sub MarkCountUpload()
    MarkAxisCountDisplay.Caption = MarkAdjustPlc.ReadWord(TAdjustReadMap.AxisCount)
End Sub
Private Function MarkCountChange() As Boolean
    If MarkAxisCountDisplay.Caption <> MarkAdjustPlc.ReadWord(TAdjustReadMap.AxisCount) Then
        MarkCountChange = True
    Else
        MarkCountChange = False
    End If
End Function


'****************************************************
'  funzioni per la gestione delle azioni dell' utente
'****************************************************
Private Sub SendFacerAxisCommand_Click()
    Dim FirstDiameter As Integer
    Dim LastDiameter As Integer
    Dim i As Integer
    
    Dim PreviusPos As Long
    
    PreviusPos = FacerAxis.Recordset.AbsolutePosition
    FacerAxis.Recordset.MoveFirst
    FirstDiameter = FacerAxis.Recordset.Fields("Diameter")
    
    FacerAxis.Recordset.MoveLast
    LastDiameter = FacerAxis.Recordset.Fields("Diameter")
    
    FacerAxis.Recordset.MoveFirst
    For i = FirstDiameter To LastDiameter
        If Not FacerAxis.Recordset.EOF Then
            FacerAdjustPlc.WriteWord(i - FirstDiameter) = FacerAxis.Recordset.Fields("Count")
            FacerAxis.Recordset.MoveNext
        End If
    Next
    FacerAxis.Recordset.AbsolutePosition = PreviusPos
End Sub

Private Sub SendMarkAxisCommand_Click()
    Dim FirstDiameter As Integer
    
    Dim LastDiameter As Integer
    Dim i As Integer
    Dim PreviusPos As Long
    
    PreviusPos = MarkAxis.Recordset.AbsolutePosition
    
    MarkAxis.Recordset.MoveFirst
    FirstDiameter = MarkAxis.Recordset.Fields("Diameter")
    
    MarkAxis.Recordset.MoveLast
    LastDiameter = MarkAxis.Recordset.Fields("Diameter")
    
    MarkAxis.Recordset.MoveFirst
    For i = FirstDiameter To LastDiameter
        If Not MarkAxis.Recordset.EOF Then
            MarkAdjustPlc.WriteWord(i - FirstDiameter) = MarkAxis.Recordset.Fields("Count")
            MarkAxis.Recordset.MoveNext
        End If
    Next
    
    MarkAxis.Recordset.AbsolutePosition = PreviusPos
    
    
End Sub

'****************************************************
'  funzione per il controllo cambio dati
'****************************************************
Private Function DataChange() As Boolean
    DataChange = False
    If FacerCountChange Then DataChange = True
    If MarkCountChange Then DataChange = True
End Function

'********************************************************
' aggiornamento dei dati interni con quelli letti dal PLC
'********************************************************
Private Sub DataUpload()
    FacerCountUpload
    MarkCountUpload
End Sub
