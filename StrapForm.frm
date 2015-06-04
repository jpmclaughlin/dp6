VERSION 5.00
Begin VB.Form StrapForm 
   BackColor       =   &H80000005&
   Caption         =   "Strap"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   9030
   WindowState     =   1  'Minimized
   Begin VB.Frame TubeFrame 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Dati tubo (mm)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6360
      TabIndex        =   7
      Top             =   5520
      Width           =   5535
      Begin VB.Label HeightDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label LengthLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lunghezza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label ThickLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Spessore"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label XLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   135
      End
      Begin VB.Label LengthDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label ThickDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label WidthDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label DimensionLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dim. (mm)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label DiameterLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Diam."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame ProductionFrame 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Produzione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   6375
      Begin VB.Label BundlesCounterDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label BundleTubesCounterDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label CurrentTubesLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Tubo n."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label CurrentBundleLabel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pacco n."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame PositionFrame 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000007&
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   1440
         Width           =   11775
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   11
            Left            =   10560
            TabIndex        =   31
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   10
            Left            =   9720
            TabIndex        =   30
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   8880
            TabIndex        =   29
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   8040
            TabIndex        =   28
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   7
            Left            =   7200
            TabIndex        =   27
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   6360
            TabIndex        =   26
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   5520
            TabIndex        =   25
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   4680
            TabIndex        =   24
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   3840
            TabIndex        =   23
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   3000
            TabIndex        =   22
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   2160
            TabIndex        =   21
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteDisplay 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1320
            TabIndex        =   20
            Top             =   0
            Width           =   765
         End
         Begin VB.Label QuoteLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Quote :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CommandButton F9Command 
         Caption         =   "F9 - Modifica"
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
         Left            =   9480
         TabIndex        =   17
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Position (mm)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   33
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label PosizioneDisplay 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   32
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   1
         X1              =   480
         X2              =   480
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   360
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   2
         X1              =   240
         X2              =   240
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   3
         X1              =   720
         X2              =   720
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   4
         X1              =   840
         X2              =   840
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   5
         X1              =   960
         X2              =   960
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   6
         X1              =   1200
         X2              =   1200
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   7
         X1              =   1320
         X2              =   1320
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   8
         X1              =   1440
         X2              =   1440
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   9
         X1              =   1680
         X2              =   1680
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   10
         X1              =   1800
         X2              =   1800
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line StrapLine 
         BorderWidth     =   2
         Index           =   11
         X1              =   1920
         X2              =   1920
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Label PositionDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   1
         Top             =   2400
         Width           =   975
      End
      Begin VB.Shape BundleShape 
         BackColor       =   &H80000003&
         BackStyle       =   1  'Opaque
         FillStyle       =   2  'Horizontal Line
         Height          =   615
         Left            =   240
         Top             =   3600
         Width           =   5535
      End
      Begin VB.Image Image1 
         Height          =   3015
         Left            =   120
         Picture         =   "StrapForm.frx":0000
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   11640
      End
   End
End
Attribute VB_Name = "StrapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************
' Costanti e variabili per disegno pacco
'**************************************************
Dim TwipOffset As Long    ' twips
Dim TwipCoeff As Double   ' twip/m
'*************************************************
' Fine variabili per disegno pacco
'**************************************************


'**************************************************
' La funzione di load del Form è utilizzata per il
' caricamento dei parametri e la inizializzazione
' della comunicazione con il PLC
'**************************************************
Private Sub Form_Load()
    '************************************
    ' Caricamento testi
    '************************************
    Me.Caption = Param.Text("StrapForm")
    'OrderFrame.Caption = Param.Text("OrderFrame")
    F9Command.Caption = Param.Text("Modify")
'    StrapsFrame.Caption = Param.Text("StrapsFrame")
    PositionFrame.Caption = Param.Text("PositionFrame") & Unit.mmString

    ' testi CurrentOrderFrame
'    OrderNameLabel.Caption = PrinterForm.PrintField(DisplayFirstFieldPosition).Caption
'    SetUpCodeLabel.Caption = Param.Text("SetUpCode")
'    CurrentOrderLabel.Caption = Param.Text("OrderFrame")

    ' testi dimensioni tubo
    TubeFrame.Caption = Param.Text("TubeData") & Unit.mmString
    DimensionLabel.Caption = Param.Text("TubeDimension")
    DiameterLabel.Caption = Param.Text("Diameter")
    ThickLabel.Caption = Param.Text("Thickness")
    LengthLabel.Caption = Param.Text("TubeLength")
    QuoteLabel.Caption = Param.Text("Quotes")
    ' testi frame contatori
    ProductionFrame.Caption = Param.Text("Production")
    CurrentBundleLabel.Caption = Param.Text("BundleNum")
    CurrentTubesLabel.Caption = Param.Text("BundleTubes")

End Sub

'**************************************************
' Funzione di aggiornamento video e comunicazione
' con il plc da chiamare in background
'**************************************************
Public Sub Update()


    '***************************************
    ' aggiorna dati read only (posizione pacco)
    '***************************************
'    If StateDB.DBChanged Then
'        ReadOnlyPictureUpdate
'    End If


End Sub

'**************************************************
' Funzione per impedire la chiusura del form da parte
' dell'utente
'**************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If UnloadMode <> vbFormControlMenu Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Activate()
    ' aggiornamento dati visualizzati in questa pagina
    PictureUpdate
End Sub



' modifica dati
Public Sub F9Command_Click()
'    Dim TmpOrder As OrderClass
'    Set TmpOrder = New OrderClass
'
'    TmpOrder.PlcCode = StrapOrder.PlcCode
'    OrdersForm.LoadOrderData TmpOrder
'    ' trasferisce il puntatore del pacco alla finestra di modifica dati
'    ModifyForm.GetData TmpOrder, True, 2
'    ' esegue la finestra modifica dati
'    ModifyForm.Show (vbModal)
'    If ModifyForm.Ok Then
'        ' salva su database la ricetta modificata
'        OrdersForm.SaveOrderData TmpOrder, False
'        ' trasmette la nuova ricetta a tutte le mappe plc
'        OrdersForm.SendModifiedData
'    End If
'    ' aggiorna l'immagine
'    PictureUpdate
End Sub




'********************************************************
' Inizio funzioni ausiliarie per aggiornamento immagine
'********************************************************
' display quote
Private Sub QuotesUpdate()
'    Dim i As Integer
'    For i = 0 To (MAX_STRAPS - 1)
'        If VisOrder.Item.StrapQuote(i) > 0 Then
'            QuoteDisplay(i).Visible = True
'            QuoteDisplay(i).Caption = Unit.m_To_Display_mm(VisOrder.Item.StrapQuote(i))
'        Else
'            QuoteDisplay(i).Visible = False
'        End If
'    Next i
End Sub

' Bundle position
Private Sub PositionUpdate()
'    Dim pos As Double
'    pos = StateDB.Word(StateMap.Position) / 1000#
'    If pos < 0 Then pos = 0
'    PositionDisplay.Caption = Unit.m_To_Display_mm(pos)
End Sub


Private Sub DrawUpdate()
'    Dim i As Integer
'    TwipOffset = (PositionFrame.Width / 2) - BundleShape.Width
'    If StrapOrder.Tube.Length > 0 Then
'        TwipCoeff = BundleShape.Width / StrapOrder.Tube.Length
'    Else
'        TwipCoeff = 0
'    End If
'    'BundleShape.Left = TwipOffset
'    For i = 0 To (MAX_STRAPS - 1)
'        StrapLine(i).Visible = False
'        StrapLine(i).X1 = TwipOffset + (StrapOrder.Strap.Quote(i) * TwipCoeff)
'        StrapLine(i).Y1 = BundleShape.Top
'        StrapLine(i).X2 = StrapLine(i).X1
'        StrapLine(i).Y2 = StrapLine(i).Y1 + BundleShape.Height
'        If StrapOrder.Strap.Quote(i) > 0 Then StrapLine(i).Visible = True
'    Next i
End Sub

Private Sub PositionDrawUpdate(Position As Integer)
'    Dim i As Integer
'    Dim RealPosition As Double   ' m
'    RealPosition = Position / 1000#
'    BundleShape.Visible = False
'    BundleShape.Left = TwipOffset + RealPosition * TwipCoeff
'    For i = 0 To (MAX_STRAPS - 1)
'        StrapLine(i).Visible = False
'        StrapLine(i).X1 = BundleShape.Left + BundleShape.Width - StrapOrder.Strap.Quote(i) * TwipCoeff
'        StrapLine(i).X2 = StrapLine(i).X1
'        If StrapOrder.Strap.Quote(i) > 0 Then StrapLine(i).Visible = True
'    Next i
'    BundleShape.Visible = True
End Sub

'aggiornamento display quote di reggiatura
Private Sub StrapQuotesUpdate()
'    Dim i As Integer
'    For i = 0 To (MAX_STRAPS - 1)
'        If StrapOrder.Strap.Quote(i) > 0 Then
'            QuoteDisplay(i).Caption = Unit.m_To_Display_mm(StrapOrder.Strap.Quote(i))
'            QuoteDisplay(i).Visible = True
'        Else
'            QuoteDisplay(i).Visible = False
'        End If
'    Next i
End Sub

'*************************************************************
'    Fine funzioni ausiliarie per disegno pacco
'*************************************************************


'****************************************************
' aggiornamento immagine per dati read only
'****************************************************
Private Sub ReadOnlyPictureUpdate()
'    ' numero pacchi fatti
'    BundlesCounterDisplay.Caption = StateDB.Word(StateMap.StrapBundleNum)
'    ' numero tubi nel pacco
'    BundleTubesCounterDisplay.Caption = StateDB.Word(StateMap.StrapTubesNum)
'    ' posizione pacco
'    PositionUpdate
'    PositionDrawUpdate StateDB.Word(StateMap.Position)
End Sub

'****************************************************
' aggiornamento immagine per tutti i dati
'****************************************************
Public Sub PictureUpdate()
'    ' disegno quote
'    QuotesUpdate
'    ' disegno reggie
'    DrawUpdate
'    ' display quote
'    StrapQuotesUpdate
'
'    PosizioneDisplay.Caption = Unit.m_To_Display_mm(StrapOrder.Strap.Posizione)
'
'
'    ' ******+ aggiorna dati ordine corrente su display di stato *********
''    CurrentPlcCodeDisplay.Text = VisOrder.PlcCode
''    CurrentOrderNameDisplay.Text = VisOrder.DisplayData
''    CurrentRecipeDisplay.Text = VisOrder.Item.Name
''    ' preset pacchi dell'ordine in corso
''    If VisOrder.AutomaticChange Or VisOrder.AutomaticStop Then
''        OfLabel(1).Visible = True
''        BundlesPresetDisplay.Visible = True
''        BundlesPresetDisplay.Caption = VisOrder.BundlesPreset
''    Else
''        OfLabel(1).Visible = False
''        BundlesPresetDisplay.Visible = False
''    End If
'
'
'    ' dimensioni tubo dell'ordine in corso
'    If StrapOrder.Tube.Round Then
'        DimensionLabel.Visible = False
'        DiameterLabel.Visible = True
'        HeightDisplay.Visible = False
'        XLabel.Visible = False
'    Else
'        DimensionLabel.Visible = True
'        DiameterLabel.Visible = False
'        HeightDisplay.Visible = True
'        XLabel.Visible = True
'    End If
'    HeightDisplay.Caption = Unit.m_To_Display_mm(StrapOrder.Tube.Height)
'    WidthDisplay.Caption = Unit.m_To_Display_mm(StrapOrder.Tube.Width)
'    ThickDisplay.Caption = Unit.m_To_Display_mm0(StrapOrder.Tube.Thickness)
'    LengthDisplay.Caption = Unit.m_To_Display_mm(StrapOrder.Tube.Length)
'
'
'    ' preset tubi pacco in corso
''    BundleTubesPresetDisplay.Caption = VisOrder.Item.Tubes
'    ' ******fine aggiornamento dati ordine corrente su display di stato *********
'
'    ' dati di stato
'    ReadOnlyPictureUpdate

End Sub


