VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BundlesLogForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin TabDlg.SSTab SSTab2 
      Height          =   2445
      Left            =   8040
      TabIndex        =   31
      Top             =   1080
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4313
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   1058
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mark as deleted bundle"
      TabPicture(0)   =   "BundlesLogForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.CommandButton Command1 
         Caption         =   "Delete selected row"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   1725
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2475
      Left            =   270
      TabIndex        =   22
      Top             =   1080
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   4366
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   8421504
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "BundlesLogForm.frx":001C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelData"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "XPButton1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "XPButton1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "XPButton1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4470
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1470
         Width           =   3075
      End
      Begin dp6.XPButton XPButton1 
         Height          =   825
         Index           =   4
         Left            =   1590
         TabIndex        =   24
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1455
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
         TX              =   "+"
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
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin dp6.XPButton XPButton1 
         Height          =   825
         Index           =   5
         Left            =   2910
         TabIndex        =   25
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1455
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
         TX              =   "-"
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
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin dp6.XPButton XPButton1 
         Height          =   825
         Index           =   8
         Left            =   270
         TabIndex        =   26
         Top             =   1380
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1455
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
            Size            =   36
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
      Begin VB.Label LabelData 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   990
         TabIndex        =   29
         Top             =   690
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   4530
         TabIndex        =   28
         Top             =   1020
         Width           =   2715
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "BundlesLogForm.frx":0038
         Top             =   90
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   570
         TabIndex        =   27
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Totals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2385
      Left            =   10080
      TabIndex        =   14
      Top             =   1080
      Width           =   5235
      Begin VB.Label LabelPacchi 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pacchi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   -210
         TabIndex        =   20
         Top             =   420
         Width           =   2715
      End
      Begin VB.Label LabelBarre 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   -150
         TabIndex        =   19
         Top             =   1050
         Width           =   2655
      End
      Begin VB.Label LabelPeso 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   -210
         TabIndex        =   18
         Top             =   1710
         Width           =   2715
      End
      Begin VB.Label DisplayPacchi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2610
         TabIndex        =   17
         Top             =   300
         Width           =   2550
      End
      Begin VB.Label DisplayTubi 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2610
         TabIndex        =   16
         Top             =   960
         Width           =   2550
      End
      Begin VB.Label DisplayPeso 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12/12/02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2610
         TabIndex        =   15
         Top             =   1620
         Width           =   2550
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   21
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
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Lay out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
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
         TabIndex        =   12
         Top             =   630
         Width           =   3495
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
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
         TabIndex        =   11
         Top             =   240
         Width           =   2415
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
         Caption         =   "Ricetta"
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
         Height          =   375
         Index           =   3
         Left            =   8610
         TabIndex        =   10
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2309"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
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
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pagina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
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
         TabIndex        =   8
         Top             =   690
         Width           =   2415
      End
      Begin VB.Label lblbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ordine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
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
         TabIndex        =   7
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "BundlesLogForm.frx":0342
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2235
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   2
         Left            =   3180
         TabIndex        =   13
         Top             =   150
         Width           =   8985
      End
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1245
      Index           =   6
      Left            =   14040
      TabIndex        =   1
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
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   300
      Top             =   900
   End
   Begin MSAdodcLib.Adodc AdoPacchi 
      Height          =   615
      Left            =   810
      Top             =   8910
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoPacchi"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridPacchi 
      Height          =   6840
      Left            =   270
      TabIndex        =   0
      Top             =   3540
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   12065
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   12632256
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
      _Band(0).Cols   =   5
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1275
      Index           =   7
      Left            =   14040
      TabIndex        =   2
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
      TabIndex        =   30
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   6825
      Index           =   1
      Left            =   13980
      Shape           =   4  'Rounded Rectangle
      Top             =   3540
      Width           =   1365
   End
End
Attribute VB_Name = "BundlesLogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DataFormat As String
Private DataDB As String
Private PesaOn As Boolean
Private DataInizio As Date
Public Filtro As String
Private TurnoInizio(3) As String
Private TurnoFine(3) As String
Private OneStep As Boolean
Private BundleSelected As Boolean

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub Combo1_Click()
   If Combo1.ListIndex = 0 Then
       Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
            " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) ORDER BY Contatore;"
   Else
        If Val(TurnoInizio(Combo1.ListIndex)) <= Val(TurnoFine(Combo1.ListIndex)) Then
            Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
            " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and ((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & "))"
        Else
            Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
            " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and (((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<=23) or (Hour(DataOra)>=0 And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & ")))"
        End If
    End If
    Aggiorna
End Sub

Private Sub Command1_Click()
    Dim cn As New ADODB.Connection
    Dim outstring As String
    Dim i As Integer
    
    TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
    TechPasswordForm.Show vbModal
    If TechPasswordForm.LoginSucceeded = False Then Exit Sub
    
    GridPacchi.Col = 5
    If BundleSelected = False Or Left(Trim(GridPacchi.Text), 2) = "D:" Then
       MsgBox "You must selected a right bundle before delete it"
       Exit Sub
    End If
    On Error GoTo error
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\History.mdb;Persist Security Info=False"
    cn.Execute "UPDATE Bundles SET Deleted=1,Cartellino9='D:" & Trim(GridPacchi.Text) & "' WHERE Cartellino9='" & Trim(GridPacchi.Text) & "'"
    cn.Close
    Set cn = Nothing
    
    outstring = ""
    outstring = outstring & Chr(34) & Format(Now, "mm/dd/yyyy") & " " & Format(Now, "hh:mm") & Chr(34) & ","
    outstring = outstring & Chr(34) & Format(Now, "mm/dd/yyyy") & Chr(34) & ","
    outstring = outstring & Chr(34) & Format(Now, "hh:mm") & Chr(34) & ","
    GridPacchi.Col = 1
    outstring = outstring & Chr(34) & Trim(GridPacchi.Text) & Chr(34) & ","
    For i = 1 To 10
        outstring = outstring & Chr(34) & "" & Chr(34) & ","
    Next i
    GridPacchi.Col = 2
    outstring = outstring & Chr(34) & Trim(GridPacchi.Text) & Chr(34) & ","
    GridPacchi.Col = 3
    outstring = outstring & Chr(34) & Trim(GridPacchi.Text) & Chr(34) & ","
    outstring = outstring & Chr(34) & "" & Chr(34) & ","
    outstring = outstring & Chr(34) & "" & Chr(34) & ","
    outstring = outstring & Chr(34) & "" & Chr(34) & ","
    GridPacchi.Col = 4
    outstring = outstring & Chr(34) & Trim(GridPacchi.Text) & Chr(34) & ","
    GridPacchi.Col = 5
    outstring = outstring & Chr(34) & "D:" & Trim(GridPacchi.Text) & Chr(34) & "#"
    
    Export.Increase_Exports_number
    Export.EXPORT_BundleLog outstring
    
    Aggiorna
    
    Exit Sub
    
error:
    MsgBox "Was not possible to delete the bundle", vbInformation, "DP6"
End Sub

Private Sub GridPacchi_Click()
    GridPacchi.Col = 0
    GridPacchi.ColSel = GridPacchi.Cols - 1
    BundleSelected = True
End Sub

'aggiornamento della pagina cortrente

Private Sub TimerLocale_Timer()
    Static NuovoCod402 As Variant
    Static VecchioCod402 As Variant
    Static NuovoNumpacco As Variant
    Static VecchioNumpacco As Variant
    Static NuovoPeso As Variant
    Static VecchioPeso As Variant
    Static NuovoNumTubi As Variant
    Static VecchioNumTubi As Variant
    Dim Risult As Boolean
    
    'aggiorna i dati pagina
    lblbar(2) = PaginaArchivia.Ordine_Descrizione
    lblbar(4) = PaginaArchivia.Ricetta_Descrizione
    
    NuovoCod402 = frmKernel.CodOrdineCorrente.Cod402
    NuovoNumpacco = DB402.Word(12)
    NuovoNumTubi = DB402.Word(14)
    NuovoPeso = DB402.Word(16)
    
    Risult = NuovoCod402 <> VecchioCod402 Or NuovoNumpacco <> VecchioNumpacco Or NuovoPeso <> VecchioPeso Or NuovoNumTubi <> VecchioNumTubi
   
    If Risult Or (Param.GetBit("Par204_AttivaPesa") <> PesaOn) Then
      Call Aggiorna
      DB402.DatiCambiati = False
      PesaOn = Param.GetBit("Par204_AttivaPesa")
      VecchioCod402 = NuovoCod402
      VecchioNumpacco = NuovoNumpacco
      VecchioPeso = NuovoPeso
      VecchioNumTubi = NuovoNumTubi
   End If
   ' aggiorna lo stato del pulsante comunicazione
   If frmKernel.StatoCom Then
      If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
   Else
       If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
   End If
End Sub

Private Sub Form_Activate()
   Dim i As Integer
   'Dim cn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim StringaSql As String
   
   SSTab1.TabCaption(0) = ""
   If OneStep = False Then
         Barra21.Selezionato = 2
         On Error Resume Next
         ' abilitazione temporizzatore locale
         PesaOn = Param.GetBit("Par204_AttivaPesa")
         TimerLocale.Enabled = True
         TimerLocale.Interval = 500
         Combo1.Clear
         Combo1.AddItem Param.Text("Tutti")
         DisplayPeso.Visible = Param.GetBit("Par204_AttivaPesa")
         LabelPeso.Visible = Param.GetBit("Par204_AttivaPesa")
         'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
         With rs
             StringaSql = "SELECT * FROM Turni"
             .Open StringaSql, Connessione, adOpenForwardOnly, adLockReadOnly, adCmdText
             For i = 1 To 3
                  TurnoInizio(i) = .Fields("TurnoInizio")
                  TurnoFine(i) = .Fields("TurnoFine")
                  Combo1.AddItem .Fields("TurnoAlias")
                 .MoveNext
             Next
             Combo1.ListIndex = 0
             .Close
             Set .ActiveConnection = Nothing
         End With
         Set rs = Nothing
         'Set cn = Nothing
         
         ' inizializzazione data
         DataInizio = Now
         DataFormat = InfoLocali(&H1F)
         DataDB = Format(Month(Now), "00") & "/" & Format(Day(Now), "00") & "/" & Format(year(Now), "0000")
         Calendario.Day = Day(DataInizio)
         Calendario.Month = Month(DataInizio)
         Calendario.year = year(DataInizio)
         
         '============================================================================================
         
         If Combo1.ListIndex = 0 Then
             Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
                        " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) ORDER BY Contatore;"
         Else
               If Val(TurnoInizio(Combo1.ListIndex)) <= Val(TurnoFine(Combo1.ListIndex)) Then
                  Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
                              " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and ((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & "))"
               Else
                  Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
                             " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and (((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<=23) or (Hour(DataOra)>=0 And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & ")))"
               End If
        End If
        Aggiorna
        OneStep = True
   End If
End Sub

Sub Aggiorna()
    Dim StringaSql As String
    Dim IndiceRiga As Integer
    Dim TotaleTubi As Long
    Dim TotalePeso As Long
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
  
    TotaleTubi = 0
    TotalePeso = 0
    LabelData.caption = DataInizio
    
    On Error Resume Next
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\History.mdb;Persist Security Info=False"

     With rs
        StringaSql = Filtro
        .Open StringaSql, cn, , adLockReadOnly, adCmdText
        If .EOF = False Then
            .MoveFirst
        End If
            
        If Param.GetBit("Par204_AttivaPesa") Then
           GridPacchi.Cols = 6
        Else
           GridPacchi.Cols = 5
        End If
        GridPacchi.Row = 0
        GridPacchi.Col = 0
        GridPacchi.Text = Param.Text("Ora")
        GridPacchi.CellAlignment = flexAlignCenterCenter
        GridPacchi.Col = 1
        GridPacchi.Text = Param.Text("Descrizione")
        GridPacchi.CellAlignment = flexAlignCenterCenter
        GridPacchi.Col = 2
        GridPacchi.Text = Param.Text("Npacco")
        GridPacchi.CellAlignment = flexAlignCenterCenter
        GridPacchi.Col = 3
        GridPacchi.Text = "Tubes"
        GridPacchi.CellAlignment = flexAlignCenterCenter
        GridPacchi.Col = 5
        GridPacchi.Text = Param.Text("Bundle")
        GridPacchi.CellAlignment = flexAlignCenterCenter
        If Param.GetBit("Par204_AttivaPesa") Then
            GridPacchi.Col = 4
            GridPacchi.CellAlignment = flexAlignCenterCenter
            If Not Param.GetBit("Par101_MisureMetriche") Then
               GridPacchi.Text = Param.Text("Peso") & " [lb]"
            Else
               GridPacchi.Text = Param.Text("Peso") & " [Kg]"
            End If
        End If
        ' set larghezza colonne
        GridPacchi.ColWidth(0) = 1000   ' ora
        GridPacchi.ColWidth(1) = 5000   ' descrizione
        GridPacchi.ColWidth(2) = 2000   ' numero pacco
        GridPacchi.ColWidth(3) = 1500   ' numero tubi
        GridPacchi.ColWidth(5) = 2000   ' numero tubi
        If Param.GetBit("Par204_AttivaPesa") Then
           GridPacchi.ColWidth(4) = 1800   ' peso pacco
        End If
        
        IndiceRiga = 1
        While .EOF = False
            If GridPacchi.Rows < (IndiceRiga + 1) Then GridPacchi.AddItem ""
            
            GridPacchi.Row = IndiceRiga
            IndiceRiga = IndiceRiga + 1
            GridPacchi.Col = 0
            GridPacchi.CellAlignment = flexAlignCenterCenter
            GridPacchi.Text = .Fields("Ora")
            
            GridPacchi.Col = 1
            GridPacchi.CellAlignment = flexAlignLeftCenter
            GridPacchi.Text = .Fields("OrdineDescrizione")
            
            GridPacchi.Col = 2
            GridPacchi.CellAlignment = flexAlignCenterCenter
            GridPacchi.Text = .Fields("NumeroPacco")
            
            GridPacchi.Col = 3
            GridPacchi.CellAlignment = flexAlignCenterCenter
            GridPacchi.Text = .Fields("NumeroTubi")
            TotaleTubi = TotaleTubi + Val(GridPacchi.Text)
            
            GridPacchi.Col = 5
            GridPacchi.CellAlignment = flexAlignCenterCenter
            GridPacchi.Text = .Fields("Cartellino9")
            If Left(Trim(GridPacchi.Text), 2) = "D:" Then
               GridPacchi.CellBackColor = vbRed
            Else
               GridPacchi.CellBackColor = vbWhite
            End If
            
            If Param.GetBit("Par204_AttivaPesa") Then
                  GridPacchi.Col = 4
                  GridPacchi.CellAlignment = flexAlignCenterCenter
                  GridPacchi.Text = Round(.Fields("PesoPacco"), 1)
                  TotalePeso = TotalePeso + Val(GridPacchi.Text)
            End If
            .MoveNext
        Wend
        .Close
        Set .ActiveConnection = Nothing
    End With
    GridPacchi.Rows = IndiceRiga
    If IndiceRiga > 1 Then GridPacchi.FixedRows = 1
    ' label totali
    DisplayPacchi.caption = CStr(GridPacchi.Rows - 1)
    DisplayPeso.caption = CStr(TotalePeso)
    DisplayTubi.caption = CStr(TotaleTubi)

Errore:

    Set rs = Nothing
    Set cn = Nothing
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
   OneStep = False
End Sub

Private Sub Form_Load()

   ScritteMultilingua
     
  'posizione data e ora
'   LblData.Top = 240
'   LblOra.Top = 480
'   LblOra.Left = 12480
'   LblData.Left = 12000
   TimerLocale.Enabled = False
   
'   Picture = LoadPicture("..\bitmap\SfondoDP6_ver2_0.jpg")
   WindowState = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormControlMenu Then
        On Error Resume Next
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

' aggiornamento dati della pagina corrente

Sub ScritteMultilingua()
   LabelPacchi.caption = Param.Text("Pacchi")
   LabelBarre.caption = Param.Text("Barre")
   Label1.caption = Param.Text("Turno")
   lblbar(3) = Param.Text("Ricette")
   lblbar(0) = Param.Text("ORDER")
   If Param.GetBit("Par101_MisureMetriche") Then
      LabelPeso.caption = Param.Text("Peso") & " [lb]"
   Else
      LabelPeso.caption = Param.Text("Peso") & " [Kg]"
   End If
   lblbar(1) = Param.Text("Pagina")
   lblbar(5) = Param.Text("Bundle log page")
   Label2 = Param.Text("Filtro")
End Sub

Private Sub XPButton1_Click(Index As Integer)
 Dim Indice As Integer
 Dim a, b, c
 
 Select Case Index
     Case 0
            frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
            frmStatistica.Show
            frmConnecting.Hide
     Case 1
           frmConnecting.ShowConnecting "Refreshing param grid. Please wait...", Me
           Param.ChiamataService = True
           Param.One = False
           frmKernel.PaginaCorrente = PagService
           frmConnecting.Hide
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
               .NomeFile = "pacchi_pagina.htm"
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .webPreview.Move 0, 0, frmHelp.Width, frmHelp.Height
               .Show
           End With
     Case 4
            DataInizio = DataInizio + 1
            a = Day(DataInizio)
            b = Month(DataInizio)
            c = year(DataInizio)
            DataDB = Format(b, "00") & "/" & Format(a, "00") & "/" & Format(c, "0000")
            If Combo1.ListIndex = 0 Then
                Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) ORDER BY Contatore;"
            Else
                 If Val(TurnoInizio(Combo1.ListIndex)) <= Val(TurnoFine(Combo1.ListIndex)) Then
                     Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and ((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & "))"
                 Else
                     Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and (((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<=23) or (Hour(DataOra)>=0 And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & ")))"
                 End If
             End If
             Aggiorna
            'Calendario.Day = Day(DataInizio)
            'Calendario.Month = Month(DataInizio)
            'Calendario.year = year(DataInizio)
      Case 5
            DataInizio = DataInizio - 1
            a = Day(DataInizio)
            b = Month(DataInizio)
            c = year(DataInizio)
            DataDB = Format(b, "00") & "/" & Format(a, "00") & "/" & Format(c, "0000")
            If Combo1.ListIndex = 0 Then
                Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) ORDER BY Contatore;"
            Else
                 If Val(TurnoInizio(Combo1.ListIndex)) <= Val(TurnoFine(Combo1.ListIndex)) Then
                     Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and ((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & "))"
                 Else
                     Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and (((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<=23) or (Hour(DataOra)>=0 And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & ")))"
                 End If
             End If
             Aggiorna
            'Calendario.Day = Day(DataInizio)
            'Calendario.Month = Month(DataInizio)
            'Calendario.year = year(DataInizio)
      Case 6
            Indice = GridPacchi.TopRow
            Indice = Indice - 10
            If Indice < 1 Then Indice = 1
            On Error Resume Next
                GridPacchi.TopRow = Indice
            On Error GoTo 0
            Aggiorna
    Case 7
            Indice = GridPacchi.TopRow
            Indice = Indice + 10
            If Indice >= GridPacchi.Rows Then Indice = GridPacchi.Rows - 1
            On Error Resume Next
                GridPacchi.TopRow = Indice
            On Error GoTo 0
            Aggiorna
    Case 8
            Calendario.year = year(DataInizio)
            Calendario.Day = Day(DataInizio)
            Calendario.Month = Month(DataInizio)
            frmcalendar.Show vbModal
            OneStep = True
            If Left(DataFormat, 2) = "dd" Then
               DataInizio = Format(Calendario.Day & " " & Calendario.Month & " " & Calendario.year, DataFormat)
            Else
               DataInizio = Format(Calendario.Month & " " & Calendario.Day & " " & Calendario.year, DataFormat)
            End If
            DataDB = Format(Calendario.Month, "00") & "/" & Format(Calendario.Day, "00") & "/" & Format(Calendario.year, "0000")
            If Combo1.ListIndex = 0 Then
                Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) ORDER BY Contatore;"
            Else
                 If Val(TurnoInizio(Combo1.ListIndex)) <= Val(TurnoFine(Combo1.ListIndex)) Then
                     Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and ((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & "))"
                 Else
                     Filtro = "SELECT * FROM Bundles WHERE  ((DataOra >= '" & DataDB & _
                     " 00.00.00' ) AND (DataOra <= '" & DataDB & " 23.59.59' )) and (((Hour(DataOra)>=" & Val(TurnoInizio(Combo1.ListIndex)) & " And Hour(DataOra)<=23) or (Hour(DataOra)>=0 And Hour(DataOra)<" & Val(TurnoFine(Combo1.ListIndex)) & ")))"
                 End If
             End If
             Aggiorna
    End Select
  Exit Sub
ErrorePercorso:
           MsgBox "Percorso file guida errato", vbExclamation
End Sub


