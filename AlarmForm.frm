VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form AlarmForm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "    Dim PosizioneAttuale As Integer"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.CheckBox CheckDB426 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB426"
      Height          =   525
      Left            =   3870
      TabIndex        =   34
      Top             =   2130
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB425"
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB425 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB425"
      Height          =   495
      Left            =   2490
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB419 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB419"
      Height          =   495
      Left            =   990
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB418 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB418"
      Height          =   495
      Left            =   5250
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB417 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB417"
      Height          =   495
      Left            =   3810
      TabIndex        =   17
      Top             =   1500
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB416 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB416"
      Height          =   495
      Left            =   2370
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB415 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB415"
      Height          =   495
      Left            =   990
      TabIndex        =   15
      Top             =   1380
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB414 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB414"
      Height          =   495
      Left            =   5250
      TabIndex        =   14
      Top             =   810
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB413 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB413"
      Height          =   495
      Left            =   3810
      TabIndex        =   13
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB412 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB412"
      Height          =   495
      Left            =   2370
      TabIndex        =   12
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB411 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB411"
      Height          =   495
      Left            =   990
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB410 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB410"
      Height          =   495
      Left            =   6720
      TabIndex        =   10
      Top             =   810
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5640
      Top             =   90
   End
   Begin VB.CheckBox CheckDB422 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB422"
      Height          =   495
      Left            =   2730
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB420 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB420"
      Height          =   495
      Left            =   1350
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB424 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB424"
      Height          =   495
      Left            =   4110
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckDB400 
      BackColor       =   &H00FF00FF&
      Caption         =   "DB400"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAllarmiPlc 
      Height          =   4455
      Left            =   30
      TabIndex        =   4
      Top             =   690
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   255
      ForeColor       =   16777215
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      WordWrap        =   -1  'True
      HighLight       =   0
      ScrollBars      =   2
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridMessaggiPlc 
      Height          =   4455
      Left            =   0
      TabIndex        =   5
      Top             =   5910
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   65535
      ForeColor       =   0
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      WordWrap        =   -1  'True
      HighLight       =   0
      ScrollBars      =   2
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1245
      Index           =   1
      Left            =   14010
      TabIndex        =   21
      Top             =   900
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2196
      TxtText         =   " "
      TxtTop          =   5
      TxtLeft         =   5
      BTYPE           =   3
      IMGTOP          =   15
      IMGLEFT         =   18
      ICONA           =   "..\Bitmap\Icone\ARW03UP.ICO"
      ImgW            =   50
      ImgH            =   50
      ImgAllarga      =   -1  'True
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      FCOL            =   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1275
      Index           =   2
      Left            =   14010
      TabIndex        =   22
      Top             =   3660
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2249
      TxtText         =   " "
      TxtTop          =   5
      TxtLeft         =   5
      BTYPE           =   3
      IMGTOP          =   15
      IMGLEFT         =   18
      ICONA           =   "..\Bitmap\Icone\ARW03DN.ICO"
      ImgW            =   50
      ImgH            =   50
      ImgAllarga      =   -1  'True
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      FCOL            =   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1245
      Index           =   3
      Left            =   14010
      TabIndex        =   23
      Top             =   6150
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2196
      TxtText         =   " "
      TxtTop          =   5
      TxtLeft         =   5
      BTYPE           =   3
      IMGTOP          =   15
      IMGLEFT         =   18
      ICONA           =   "..\Bitmap\Icone\ARW03UP.ICO"
      ImgW            =   50
      ImgH            =   50
      ImgAllarga      =   -1  'True
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      FCOL            =   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1275
      Index           =   4
      Left            =   14010
      TabIndex        =   24
      Top             =   8910
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   2249
      TxtText         =   " "
      TxtTop          =   5
      TxtLeft         =   5
      BTYPE           =   3
      IMGTOP          =   15
      IMGLEFT         =   18
      ICONA           =   "..\Bitmap\Icone\ARW03DN.ICO"
      ImgW            =   50
      ImgH            =   50
      ImgAllarga      =   -1  'True
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   65535
      FCOL            =   0
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   25
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Frame FrameErrHelp 
      BackColor       =   &H00C0C0C0&
      Height          =   11475
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   15315
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   5
         Left            =   90
         TabIndex        =   32
         Top             =   10530
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1561
         TxtText         =   "Alarm info"
         TxtTop          =   35
         TxtLeft         =   50
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\MSGBOX02.ICO"
         ImgW            =   50
         ImgH            =   20
         ImgAllarga      =   0   'False
         TX              =   "      "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   0
         Left            =   12990
         TabIndex        =   31
         Top             =   10530
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1561
         TxtText         =   "Alarm list"
         TxtTop          =   35
         TxtLeft         =   50
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\prjdll15.ico"
         ImgW            =   50
         ImgH            =   20
         ImgAllarga      =   0   'False
         TX              =   "      "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   6
         Left            =   2400
         TabIndex        =   30
         Top             =   10530
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1561
         TxtText         =   "Test I/O"
         TxtTop          =   35
         TxtLeft         =   50
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\enet1.ico"
         ImgW            =   50
         ImgH            =   20
         ImgAllarga      =   0   'False
         TX              =   "      "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   7
         Left            =   90
         TabIndex        =   29
         Top             =   10530
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1561
         TxtText         =   "Electric sheets"
         TxtTop          =   35
         TxtLeft         =   30
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\MEMacroEditor0.ico"
         ImgW            =   50
         ImgH            =   20
         ImgAllarga      =   0   'False
         TX              =   "    "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin SHDocVwCtl.WebBrowser webPreview 
         Height          =   10335
         Left            =   60
         TabIndex        =   27
         Top             =   150
         Width           =   15165
         ExtentX         =   26749
         ExtentY         =   18230
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
         Location        =   "http:///"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alarm help file not found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   840
         Left            =   2880
         TabIndex        =   28
         Top             =   4140
         Width           =   8400
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   4425
      Index           =   0
      Left            =   13950
      Shape           =   4  'Rounded Rectangle
      Top             =   5970
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   4425
      Index           =   1
      Left            =   13950
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label LabelAllarmi 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "ALLARMI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Index           =   0
      Left            =   5490
      TabIndex        =   8
      Top             =   0
      Width           =   4305
   End
   Begin VB.Label LabelMessaggi 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGGI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   675
      Index           =   1
      Left            =   5220
      TabIndex        =   6
      Top             =   5220
      Width           =   4785
   End
   Begin VB.Label LabelMessaggi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MESSAGGI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   5190
      Width           =   15330
   End
   Begin VB.Label LabelAllarmi 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ALLARMI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15345
   End
End
Attribute VB_Name = "AlarmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MaxAllarmiPerZona As Integer = 16

Private OneStep As Boolean
Private NumAllOld  As Long
Private NumMsgold As Long
Private PathFile As String

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   OneStep = False
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub Barra21_RipetizioneTasto()
   frmKernel.PulAllarmiPremuto = False
   frmKernel.PaginaCorrente = 7
   AggiornaListaAllarmi
   AggiornaListaMessaggi
   frmKernel.PulAllarmiPremuto = False
End Sub

Private Sub Form_Activate()
    If OneStep Then Exit Sub
    Barra21.Selezionato = 7
    Timer1.Enabled = True
    Timer1.Interval = 500
    AggiornaListaAllarmi
    AggiornaListaMessaggi
    frmKernel.PulAllarmiPremuto = False
    FrameErrHelp.Visible = False
    OneStep = True
End Sub

Private Sub Form_Deactivate()
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
   ScritteMultilingua
End Sub

Private Sub GridAllarmiPlc_Click()
'   Dim src
'   Dim PathFile As String
'   Dim AddIndexs() As String
'   Dim AlFile As String
'
'   On Error Resume Next
'   webPreview.Visible = False
'   src = Left(GridAllarmiPlc.Text, InStr(GridAllarmiPlc.Text, ":") - 1)
'   AddIndexs() = Split(src, "_")
'   AlFile = frmKernel.HelpAllarme(AddIndexs(0), CInt(AddIndexs(1)), CInt(AddIndexs(2))) ' Left(GridAllarmiPlc.Text, src - 2) & ".htm"
'   If AlFile = "NULL" Then Exit Sub
'   If FileEsistente(PathFile) Then
'      webPreview.Navigate PathFile
'      webPreview.Visible = True
'   End If
'   FrameErrHelp.Visible = True
'   FrameErrHelp.ZOrder
'   XPButton1(7).Visible = True
'   XPButton1(5).Visible = False
End Sub

Private Sub GridMessaggiPlc_Click()
'   Dim src
'   Dim PathFile As String
'   Dim AddIndexs() As String
'   Dim AlFile As String
'
'   On Error Resume Next
'   webPreview.Visible = False
'   src = Left(GridMessaggiPlc.Text, InStr(GridMessaggiPlc.Text, ":") - 1)
'   AddIndexs() = Split(src, "_")
'   AlFile = frmKernel.HelpAllarme(AddIndexs(0), CInt(AddIndexs(1)), CInt(AddIndexs(2))) ' Left(GridAllarmiPlc.Text, src - 2) & ".htm"
'   If AlFile = "NULL" Then Exit Sub
'   PathFile = HelpPath & "Alarms\" & AlFile
'   If FileEsistente(PathFile) Then
'      webPreview.Navigate PathFile
'      webPreview.Visible = True
'   End If
'   FrameErrHelp.Visible = True
'   FrameErrHelp.ZOrder
End Sub

Private Sub Timer1_Timer()
   If (NumAllOld <> frmKernel.NumAllAttivi) Then AggiornaListaAllarmi: NumAllOld = frmKernel.NumAllAttivi
   If (NumMsgold <> frmKernel.NumMsgAttivi) Then AggiornaListaMessaggi: NumMsgold = frmKernel.NumMsgAttivi
End Sub

Private Sub AggiornaListaAllarmi()
    Dim IndiceRiga As Integer
    IndiceRiga = 0
    InsertItems GridAllarmiPlc, ListaAllarmiPC, IndiceRiga
    If (frmKernel.PcAlarm(PlcCommErr) = False) And (frmKernel.PcAlarm(PlcDDEFault) = False) Then
        If CheckDB400.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB400, IndiceRiga
        If CheckDB410.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB410, IndiceRiga
        If CheckDB411.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB411, IndiceRiga
        If CheckDB412.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB412, IndiceRiga
        If CheckDB413.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB413, IndiceRiga
        If CheckDB414.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB414, IndiceRiga
        If CheckDB415.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB415, IndiceRiga
        If CheckDB416.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB416, IndiceRiga
        If CheckDB417.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB417, IndiceRiga
        If CheckDB418.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB418, IndiceRiga
        If CheckDB419.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB419, IndiceRiga
        If CheckDB420.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB420, IndiceRiga
        If CheckDB422.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB422, IndiceRiga
        If CheckDB424.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB424, IndiceRiga
        If CheckDB425.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB425, IndiceRiga
        If CheckDB426.value = 1 Then InsertItems GridAllarmiPlc, ListaAllarmiDB426, IndiceRiga
    End If
    GridAllarmiPlc.Rows = IndiceRiga
End Sub
Private Sub AggiornaListaMessaggi()
    Dim IndiceRiga As Integer
    IndiceRiga = 0
    If (frmKernel.PcAlarm(PlcCommErr) = False) And (frmKernel.PcAlarm(PlcDDEFault) = False) Then
        If CheckDB400.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB400, IndiceRiga
        If CheckDB410.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB410, IndiceRiga
        If CheckDB411.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB411, IndiceRiga
        If CheckDB412.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB412, IndiceRiga
        If CheckDB413.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB413, IndiceRiga
        If CheckDB414.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB414, IndiceRiga
        If CheckDB415.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB415, IndiceRiga
        If CheckDB416.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB416, IndiceRiga
        If CheckDB417.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB417, IndiceRiga
        If CheckDB418.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB418, IndiceRiga
        If CheckDB419.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB419, IndiceRiga
        If CheckDB420.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB420, IndiceRiga
        If CheckDB422.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB422, IndiceRiga
        If CheckDB424.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB424, IndiceRiga
        If CheckDB425.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB425, IndiceRiga
        If CheckDB426.value = 1 Then InsertItems GridMessaggiPlc, ListaMessaggiDB426, IndiceRiga
    End If
    GridMessaggiPlc.Rows = IndiceRiga
End Sub


Private Sub InsertItems(GridTesti As MSHFlexGrid, ListaAllarmiDB() As String, ByRef IndiceRiga As Integer)
    Dim IndiceAllarme As Integer
    GridTesti.ColWidth(0) = GridTesti.Width
    IndiceAllarme = 0
    While ListaAllarmiDB(IndiceAllarme) <> "" And IndiceAllarme < MaxAllarmiPerZona
        If GridTesti.Rows < (IndiceRiga + 1) Then GridTesti.AddItem ""
        GridTesti.Row = IndiceRiga
        GridTesti.RowHeight(IndiceRiga) = 900
        GridTesti.CellAlignment = flexAlignLeftCenter
        GridTesti.Text = ListaAllarmiDB(IndiceAllarme)
        IndiceAllarme = IndiceAllarme + 1
        IndiceRiga = IndiceRiga + 1
        DoEvents
    Wend
End Sub

Sub ScritteMultilingua()
    LabelAllarmi(0).caption = Param.Text("Allarmi")
    LabelAllarmi(1).caption = Param.Text("Allarmi")
    LabelMessaggi(0).caption = Param.Text("Messaggi")
    LabelMessaggi(1).caption = Param.Text("Messaggi")
End Sub

Private Sub XPButton1_Click(Index As Integer)
    Dim PosizioneAttuale As Integer
    
   Select Case Index
   Case 0
            FrameErrHelp.Visible = False
   Case 1
            PosizioneAttuale = GridAllarmiPlc.TopRow
            PosizioneAttuale = PosizioneAttuale - 3
            If PosizioneAttuale < 0 Then PosizioneAttuale = 0
            If GridAllarmiPlc.Rows > 0 Then GridAllarmiPlc.TopRow = PosizioneAttuale
   Case 2
            PosizioneAttuale = GridAllarmiPlc.TopRow
            PosizioneAttuale = PosizioneAttuale + 3
            If PosizioneAttuale >= GridAllarmiPlc.Rows Then PosizioneAttuale = GridAllarmiPlc.Rows - 1
            If GridAllarmiPlc.Rows > 0 Then GridAllarmiPlc.TopRow = PosizioneAttuale
   Case 3
            PosizioneAttuale = GridMessaggiPlc.TopRow
            PosizioneAttuale = PosizioneAttuale - 3
            If PosizioneAttuale < 0 Then PosizioneAttuale = 0
            If GridMessaggiPlc.Rows > 0 Then GridMessaggiPlc.TopRow = PosizioneAttuale
   Case 4
            PosizioneAttuale = GridMessaggiPlc.TopRow
            PosizioneAttuale = PosizioneAttuale + 3
            If PosizioneAttuale >= GridMessaggiPlc.Rows Then PosizioneAttuale = GridMessaggiPlc.Rows - 1
            If GridMessaggiPlc.Rows > 0 Then GridMessaggiPlc.TopRow = PosizioneAttuale
   Case 5
            '==========================
            On Error Resume Next
            XPButton1(7).Visible = True
            XPButton1(5).Visible = False
            webPreview.Visible = False
            If FileEsistente(PathFile) Then
               webPreview.Navigate PathFile
               webPreview.Visible = True
            End If
            '==========================
   Case 6
            On Error Resume Next
            Unload frmHelp
            Unload frmTestIO
            Set frmHelp = Nothing
            Set frmTestIO = Nothing
            With frmTestIO
                .Top = 0
                .Left = 0
                .WindowState = 2
                .Show
            End With
   Case 7
            On Error Resume Next
            XPButton1(7).Visible = False
            XPButton1(5).Visible = True
            webPreview.Visible = False
            If FileEsistente(HelpPath & "2202a.pdf") Then
               webPreview.Navigate HelpPath & "2202a.pdf"
               webPreview.Visible = True
            End If
   End Select
   Exit Sub
   
End Sub
