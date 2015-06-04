VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmAnalisiAllarmi 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2445
      Left            =   270
      TabIndex        =   17
      Top             =   1110
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   4313
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAnalisiAllarmi.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelValore"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelData"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Image1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "AdoAllarmiStat"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "XPButton1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "List1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdApplica"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.CommandButton CmdApplica 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Applica filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   11520
         Picture         =   "frmAnalisiAllarmi.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   450
         Width           =   1725
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   1650
         TabIndex        =   20
         Top             =   240
         Width           =   3495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Desc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   690
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Asc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   1395
      End
      Begin dp6.XPButton XPButton1 
         Height          =   555
         Index           =   8
         Left            =   9210
         TabIndex        =   21
         Top             =   630
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         TxtText         =   " "
         TxtTop          =   5
         TxtLeft         =   5
         BTYPE           =   4
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   ""
         ImgW            =   10
         ImgH            =   10
         ImgAllarga      =   0   'False
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   27.75
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
      Begin MSAdodcLib.Adodc AdoAllarmiStat 
         Height          =   495
         Left            =   60
         Top             =   1800
         Visible         =   0   'False
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   873
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"frmAnalisiAllarmi.frx":045E
         OLEDBString     =   $"frmAnalisiAllarmi.frx":04ED
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "AllarmiMultiLingua"
         Caption         =   "AdoAllarmiStat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   0
         Picture         =   "frmAnalisiAllarmi.frx":057C
         Top             =   30
         Width           =   480
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   510
         TabIndex        =   25
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valore"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   6150
         TabIndex        =   24
         Top             =   300
         Width           =   2505
      End
      Begin VB.Label LabelValore 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5700
         TabIndex        =   23
         Top             =   630
         Width           =   3465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7380
         TabIndex        =   22
         Top             =   1830
         Width           =   165
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   8790
      Top             =   2400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   5
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Test I/O"
         TxtTop          =   35
         TxtLeft         =   35
         BTYPE           =   3
         IMGTOP          =   -5
         IMGLEFT         =   -10
         ICONA           =   "..\bitmap\semaforoVerde.gif"
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
         Index           =   1
         Left            =   1650
         TabIndex        =   6
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Service"
         TxtTop          =   35
         TxtLeft         =   35
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\explorer10.ico"
         ImgW            =   10
         ImgH            =   10
         ImgAllarga      =   0   'False
         TX              =   ""
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
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Statistics"
         TxtTop          =   35
         TxtLeft         =   20
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\histdata0.ico"
         ImgW            =   10
         ImgH            =   10
         ImgAllarga      =   0   'False
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Index           =   3
         Left            =   13800
         TabIndex        =   16
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "           Help"
         TxtTop          =   35
         TxtLeft         =   10
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   4
         ICONA           =   "..\Bitmap\W95MBX04.ICO"
         ImgW            =   10
         ImgH            =   10
         ImgAllarga      =   0   'False
         TX              =   "     "
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
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "frmAnalisiAllarmi.frx":0886
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2235
      End
      Begin VB.Label lblbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   5370
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pagina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   5370
         TabIndex        =   12
         Top             =   690
         Width           =   2415
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2309"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   465
         Index           =   2
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Ricetta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   8610
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Line Line3 
         X1              =   5430
         X2              =   12060
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   4
         Left            =   9720
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Lay out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Index           =   5
         Left            =   8610
         TabIndex        =   8
         Top             =   630
         Width           =   3495
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   2
         Left            =   3180
         TabIndex        =   14
         Top             =   150
         Width           =   8985
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridParametri 
      Bindings        =   "frmAnalisiAllarmi.frx":2914
      Height          =   6840
      Left            =   270
      TabIndex        =   0
      Top             =   3540
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   12065
      _Version        =   393216
      Rows            =   5
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   12632256
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1245
      Index           =   6
      Left            =   14040
      TabIndex        =   2
      Top             =   3720
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
      Index           =   7
      Left            =   14040
      TabIndex        =   3
      Top             =   8970
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
      TabIndex        =   15
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   6855
      Index           =   6
      Left            =   13980
      Shape           =   4  'Rounded Rectangle
      Top             =   3540
      Width           =   1365
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   11430
      TabIndex        =   1
      Top             =   3060
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "frmAnalisiAllarmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Filtro As String
Private nTurno As Integer
Private FiltroApplicato As String
Private RS_Analisi As ADODB.Recordset
Private Ordinamento As String
Private Lancio As Boolean
Private AllarmeLingua As String

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
      frmKernel.PaginaCorrente = Index
      Unload Me
End Sub

Private Sub CmdApplica_Click()
    On Error Resume Next
    frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
    RefreshTabella FiltroApplicato
    Unload frmConnecting
End Sub

Private Sub Command1_Click()
   Hide
   Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
     
    'On Error Resume Next
    AdoAllarmiStat.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\AllarmiMultiLingua.mdb;Persist Security Info=False"
    AdoAllarmiStat.RecordSource = "AllarmiMultiLingua"
    AdoAllarmiStat.Refresh
    AdoAllarmiStat.Recordset.ActiveConnection = Nothing
    DoEvents
    Select Case Param.GetNumber("Par100_Lingua")
        Case 1
            AllarmeLingua = "ITALIANO"
        Case 2
            AllarmeLingua = "INGLESE"
        Case 3
            AllarmeLingua = "FRANCESE"
        Case 4
            AllarmeLingua = "SPAGNOLO"
        Case 5
            AllarmeLingua = "TEDESCO"
        Case 6
            AllarmeLingua = "LinguaSpeciale"
        Case Else
            AllarmeLingua = "ITALIANO"
    End Select
    SSTab1.TabCaption(0) = ""
    Calendario.year = year(Now)
    Calendario.Day = Day(Now)
    Calendario.Month = Month(Now)
    Lancio = False
    Set RS_Analisi = Nothing
    Set RS_Analisi = New ADODB.Recordset

    
    ScritteMultilingua
    DoEvents
    WindowState = 2
    Ordinamento = "Data ASC"
    FiltroApplicato = ""
    Option2.value = True
        
    With List1
      .Clear
      .AddItem Param.Text("Tutti")
      .AddItem Param.Text("Allarme")
      .AddItem Param.Text("Turno")
      .AddItem Param.Text("Data")
      .AddItem Param.Text("Giorno")
      .AddItem Param.Text("Ora")
      .ListIndex = 0
    End With
    GridParametri.ColHeaderCaption(0, 0) = Param.Text("Data")
    GridParametri.ColHeaderCaption(0, 1) = Param.Text("Stato")
    GridParametri.ColHeaderCaption(0, 2) = Param.Text("Descrizione")
    DoEvents
    
    RefreshTabella FiltroApplicato
    ' abilita timers
    TimerLocale.Enabled = True
    TimerLocale.Interval = 250
    'Calendar1.Today
    XPButton1(8).Visible = False
   ' Barra21.Refresh_lingua
    Barra21.Selezionato = 2
    Lancio = True
End Sub
Sub RefreshTabella(ByVal Filtro As String)
   Dim test
  
   On Error Resume Next
   If AlarmLog.TestFile(TargetPath & "Alarmslog.txt", 11, 10, 5, 1, 1) Then
      Set RS_Analisi = AlarmLog.Rspopolato(TargetPath & "Alarmslog.txt", 11, 10, 5, 1, 1)
   End If
   With RS_Analisi
        If Filtro = "" Or Filtro = "0" Then
            .Filter = adFilterNone
        Else
           .Filter = Filtro
        End If
        IndiceRiga = 0
        i = 1
        GridParametri.Clear
        GridParametri.FixedRows = 1
        GridParametri.GridColorFixed = vbRed
        GridParametri.ColHeaderCaption(0, 0) = Param.Text("Data")
        GridParametri.ColHeaderCaption(0, 1) = Param.Text("Stato")
        GridParametri.ColHeaderCaption(0, 2) = Param.Text("Descrizione")
        GridParametri.ColWidth(0) = 2000
        GridParametri.ColWidth(1) = 800
        GridParametri.ColWidth(2) = 10700
        GridParametri.Font.Name = "Arial"
        GridParametri.Font.Size = 8
        GridParametri.Rows = 0
        If .RecordCount <= 0 Then Exit Sub
        GridParametri.Rows = .RecordCount - 1
        .MoveFirst
        While .EOF = False
            DoEvents
            If GridParametri.Rows <= (IndiceRiga + 1) Then GridParametri.AddItem ""
            GridParametri.Row = i
            GridParametri.RowHeight(i) = 400
            If Trim(.Fields("Data")) <> "" Then
                GridParametri.Col = 0
                GridParametri.CellAlignment = flexAlignLeftCenter
                GridParametri.Text = Trim(.Fields("Data")) & "  " & Trim(.Fields("Ora"))
                
                GridParametri.Col = 1
                GridParametri.CellAlignment = flexAlignCenterCenter
                GridParametri.Text = IIf(Trim(.Fields("Stato")) = "1", "ON", "OFF")
                
                GridParametri.Col = 2
                GridParametri.CellAlignment = flexAlignLeftCenter
                GridParametri.Text = Trim(DescrAllarme(.Fields("Descrizione")))
                i = i + 1
            End If
            IndiceRiga = IndiceRiga + 1
            .MoveNext
            Label2 = IndiceRiga
       Wend
   End With
   
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   TimerLocale.Enabled = False
   Set RS_Analisi = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RS_Analisi = Nothing
   Unload Me
End Sub

Private Sub LabelAnno_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1900
    TOUCHNumericPad.ValoreMax = 3000
    TOUCHNumericPad.Dati = LabelAnno.Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            LabelAnno.Caption = TOUCHNumericPad.Dati
        On Error GoTo 0
    End If
End Sub

Private Sub LabelMese_Click()
  TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1
    TOUCHNumericPad.ValoreMax = 12
    TOUCHNumericPad.Dati = LabelMese.Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            LabelMese.Caption = TOUCHNumericPad.Dati
        On Error GoTo 0
    End If
End Sub
Public Sub Update()
    LabelOra.Caption = Format(Now, "hh")
    LabelMinuti.Caption = Mid(Format(Now, "hh:mm"), 4, 2)
    LabelGiorno.Caption = Format(Now, "dd")
    LabelMese.Caption = Format(Now, "mm")
    LabelAnno.Caption = Format(Now, "yyyy")
    ComPesa = Param.GetNumber("Par115_SerialePesa")
    LblCOMPesa.Caption = ComPesa
End Sub

Private Sub GridParametri_Click()
  
   On Error Resume Next
   GridParametri.Col = 0
   GridParametri.ColSel = 0
   Label3 = GridParametri.Text
End Sub

Private Sub LabelValore_Change()
    Dim i As Integer
    
    On Error Resume Next
    IIf LabelValore.Caption = "", CmdApplica.Enabled = False, CmdApplica.Enabled = True
    Select Case List1.ListIndex
       Case 0
          Filtro = ""
       Case 1
          Filtro = "Descrizione LIKE '*" & LabelValore.Caption & "*'"
       Case 2
           If Val(LabelValore) >= 1 And Val(LabelValore) <= 3 Then LabelValore = DatiTurno(Val(LabelValore), 3)
           For i = 1 To 3
              If LabelValore = DatiTurno(i, 3) Then Filtro = "Turno like '" & Trim(Str(i)) & "'": Exit For
           Next
       Case 3
            Filtro = "Data like '*" & Trim(LabelValore) & "*'"
       Case 4
            Filtro = "Data like '*/" & Format(Trim(LabelValore), "00") & "*'"
       Case 5
            Filtro = "Ora LIKE '" & LabelValore & ":*'"
    End Select
    FiltroApplicato = Filtro
End Sub

Private Sub LabelValore_Click()
    TOUCHKeyBoard.Dati = LabelValore.Caption
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        LabelValore.Caption = TOUCHKeyBoard.Dati
    End If
End Sub

Private Sub List1_Click()
   On Error Resume Next
   Select Case List1.ListIndex
       Case 0
                   XPButton1(8).Visible = False
                   GridParametri.Row = 1
                   GridParametri.Col = 2
                   LabelValore.Caption = ""
                   FiltroApplicato = ""
      Case 1
                   XPButton1(8).Visible = False
                   GridParametri.Row = 1
                   GridParametri.Col = 2
                   LabelValore.Caption = Left(GridParametri.Text, 11)
                   FiltroApplicato = "Descrizione LIKE '*" & LabelValore.Caption & "*'"
       Case 2
                   XPButton1(8).Visible = False
                   GridParametri.Row = 1
                   GridParametri.Col = 2
                   LabelValore.Caption = UCase(DatiTurno(1, 3))
                   If Val(LabelValore) >= 1 And Val(LabelValore) <= 3 Then LabelValore = DatiTurno(Val(LabelValore), 3)
                   FiltroApplicato = "Turno like '1'" ' & UCase(LabelValore) & "'"
       Case 3
                   XPButton1(8).Visible = True
                   GridParametri.Row = 1
                   GridParametri.Col = 2
                   LabelValore.Caption = Format(Calendario.Month, "00") & "/" & Format(Calendario.Day, "00") & "/" & Calendario.year
                   FiltroApplicato = "Data like '*" & Trim(LabelValore) & "*'"
       Case 4
                   XPButton1(8).Visible = True
                   GridParametri.Row = 1
                   GridParametri.Col = 2
                   LabelValore.Caption = Format(Calendario.Day, "00")
                   FiltroApplicato = "Data like '*/" & Format(Trim(LabelValore), "00") & "*'"
       Case 5
                   XPButton1(8).Visible = False
                   GridParametri.Row = 1
                   GridParametri.Col = 2
                   LabelValore.Caption = "01"
                   FiltroApplicato = "Ora LIKE '" & LabelValore & ":*'"
   End Select
   
End Sub
                    
Sub ScritteMultilingua()
    lblbar(1) = Param.Text("Pagina")
    lblbar(5) = Param.Text("StoricoAllarmi")
    Label1.Caption = Param.Text("Value")
    'Label4.Caption = Param.Text("000000046")
    CmdApplica.Caption = Param.Text("ApplicaFiltro")
   ' Command1.Caption = Param.Text("Chiudi")
   ' Command3.Caption = Param.Text("GestFile")
    LabelData.Caption = Param.Text("Filtro")
    'CmdLogDel.Caption = Param.Text("CancLogAllarmi")
    'Label2 = Param.Text("Opzioni")
End Sub

Private Sub Option1_Click()
   If Lancio = False Then Exit Sub
   Ordinamento = "Data ASC"
   On Error Resume Next
   frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
    RefreshTabella FiltroApplicato
    frmConnecting.Hide
End Sub

Private Sub Option2_Click()
   If Lancio = False Then Exit Sub
   Ordinamento = "Data DESC"
   On Error Resume Next
   frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
    RefreshTabella FiltroApplicato
    frmConnecting.Hide
End Sub



Private Sub TimerLocale_Timer()
  ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    Barra21.Allarme = frmKernel.AllarmeON
End Sub

Private Sub XPButton1_Click(Index As Integer)
 Dim PosizioneAttuale As Integer
 
 On Error Resume Next
 Select Case Index
     Case 0
            frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
            frmStatistica.Show
            frmConnecting.Hide
            Unload Me
     Case 1
           frmConnecting.ShowConnecting "Refreshing param grid. Please wait...", Me
           Param.ChiamataService = True
           Param.One = False
           frmKernel.PaginaCorrente = PagService
           frmConnecting.Hide
           Set cn = Nothing
           Unload Me
     Case 2
           On Error GoTo ErrorePercorso
           Unload frmHelp
           Unload frmTestIO
           Set frmHelp = Nothing
           Set frmTestIO = Nothing
           With frmTestIO
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .Show
           End With
     Case 3
          On Error GoTo ErrorePercorso
           Unload frmHelp
           Set frmHelp = Nothing
           With frmHelp
                .Errori = True
                .NomeFile = "Alarmslog_pagina.htm"
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .webPreview.Move 0, 0, frmHelp.Width, frmHelp.Height
               .Show
           End With
     Case 4
     Case 6
        PosizioneAttuale = GridParametri.TopRow
        PosizioneAttuale = PosizioneAttuale - 10
        If PosizioneAttuale < 0 Then PosizioneAttuale = 0
        If GridParametri.Rows > 0 Then GridParametri.TopRow = PosizioneAttuale
     Case 7
        PosizioneAttuale = GridParametri.TopRow
        PosizioneAttuale = PosizioneAttuale + 10
        If PosizioneAttuale >= GridParametri.Rows Then PosizioneAttuale = GridParametri.Rows - 1
        If GridParametri.Rows > 0 Then GridParametri.TopRow = PosizioneAttuale
     Case 8
        LabelValore = ""
        Calendario.year = year(Now)
        Calendario.Day = Day(Now)
        Calendario.Month = Month(Now)
        frmcalendar.Show vbModal
      ' If LabelValore = "" Then Exit Sub
        GridParametri.Row = 1
        GridParametri.Col = 2
        If List1.ListIndex = 3 Then
           LabelValore.Caption = Format(Calendario.Day, "00") & "/" & Format(Calendario.Month, "00") & "/" & Calendario.year
           Filtro = "Data like '*" & Trim(Format(Calendario.Month, "00") & "/" & Format(Calendario.Day, "00") & "/" & Calendario.year) & "*'"
           FiltroApplicato = Filtro
        ElseIf List1.ListIndex = 4 Then
              LabelValore.Caption = Format(Calendario.Day, "00")
              Filtro = "Data like '*/" & Format(Trim(Calendario.Day), "00") & "*'"
              FiltroApplicato = Filtro
           Else
            '  LabelValore.Caption = Format(Calendario.Day, "00")
           '   FiltroApplicato = "Ora like '" & Format(Calendario.Day, "00") & "*'"
           End If
  End Select
  Exit Sub
  
  
ErrorePercorso:
           MsgBox "The path of th help file is wrong", vbExclamation, "DP6"
End Sub
Private Function DescrAllarme(DBName As String) As String
    Dim AlarmID As String
    
    AlarmID = Trim(DBName) ' & "_" & Format(IndiceByte, "000") & "_" & IndiceBit
    AdoAllarmiStat.Recordset.MoveFirst
    AdoAllarmiStat.Recordset.Find ("TagName = '" & AlarmID & "'")
    If AdoAllarmiStat.Recordset.EOF = False Then
       If (AdoAllarmiStat.Recordset.Fields("Riferimento") = "") Or IsNull(AdoAllarmiStat.Recordset.Fields("Riferimento")) Then
           TestoAllarme = AlarmID & " : " & AdoAllarmiStat.Recordset.Fields(AllarmeLingua)
       Else
           TestoAllarme = AlarmID & " : " & AdoAllarmiStat.Recordset.Fields(AllarmeLingua) & " [" & AdoAllarmiStat.Recordset.Fields("Riferimento") & "]"
       End If
    Else
        TestoAllarme = AlarmID & " : "
    End If
    DescrAllarme = TestoAllarme
End Function
