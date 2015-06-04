VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FilettoForm 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   63
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   3
         Left            =   13800
         TabIndex        =   64
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Help"
         TxtTop          =   35
         TxtLeft         =   55
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   4
         ICONA           =   "..\Bitmap\manuale.ico"
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
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   65
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Com"
         TxtTop          =   35
         TxtLeft         =   50
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "FilettoForm.frx":0000
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
         TabIndex        =   74
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   120
      Top             =   1950
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Cambio filetto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1290
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1470
      Width           =   1995
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   970
      BackColor       =   -2147483645
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Filettatrici"
      TabPicture(0)   =   "FilettoForm.frx":208E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Tappatrici"
      TabPicture(1)   =   "FilettoForm.frx":20AA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Selettore(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Selettore(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label01(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label01(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label11"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Manicottatrici"
      TabPicture(2)   =   "FilettoForm.frx":20C6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Selettore(4)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label01(4)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label15(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "SpessManicottoDisplay(0)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "NumGiriManicotto"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label23"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Verniciatrici"
      TabPicture(3)   =   "FilettoForm.frx":20E2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Selettore(5)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Selettore(6)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label01(6)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label01(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label13"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label14"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label16"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "TempoVernDisplay(0)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7275
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   570
         Width           =   7635
         Begin MSAdodcLib.Adodc AdoFiletti 
            Height          =   525
            Left            =   4440
            Top             =   4410
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   926
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
            Connect         =   $"FilettoForm.frx":20FE
            OLEDBString     =   $"FilettoForm.frx":2191
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Filetti"
            Caption         =   "AdoFiletti"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   3
            Left            =   900
            TabIndex        =   58
            Top             =   3900
            Width           =   165
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   2
            Left            =   720
            TabIndex        =   57
            Top             =   3900
            Width           =   165
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   1
            Left            =   540
            TabIndex        =   56
            Top             =   3900
            Width           =   165
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   0
            Left            =   360
            TabIndex        =   55
            Top             =   3900
            Width           =   165
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   3390
            Width           =   1215
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "RPM"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   330
            TabIndex        =   49
            Top             =   3120
            Width           =   795
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Index           =   0
            Left            =   120
            TabIndex        =   47
            Top             =   2700
            Width           =   1215
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "Passo"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Left            =   330
            TabIndex        =   45
            Top             =   2430
            Width           =   795
         End
         Begin VB.Label Label7 
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "Quota stop"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   645
            Left            =   3510
            TabIndex        =   32
            Top             =   5940
            Width           =   3255
         End
         Begin VB.Label Label6 
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "Quota start"
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
            Height          =   645
            Left            =   510
            TabIndex        =   31
            Top             =   5940
            Width           =   3255
         End
         Begin VB.Label Label01 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0      1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   7
            Left            =   6060
            TabIndex        =   24
            Top             =   5730
            Width           =   1215
         End
         Begin VB.Image Selettore 
            Height          =   1155
            Index           =   0
            Left            =   6060
            Top             =   6030
            Width           =   1185
         End
         Begin VB.Image ImgStart1 
            Height          =   4515
            Left            =   1470
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Image ImgStop1 
            Height          =   4920
            Left            =   1590
            Top             =   960
            Width           =   1110
         End
         Begin VB.Image Fil1_base 
            Height          =   2565
            Left            =   3120
            Top             =   1740
            Width           =   855
         End
         Begin VB.Label StopQuoteDisplay 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   3780
            TabIndex        =   14
            Top             =   6450
            Width           =   1455
         End
         Begin VB.Label StartQuoteDisplay 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   780
            TabIndex        =   13
            Top             =   6450
            Width           =   1455
         End
         Begin VB.Image ImgFilSx 
            Height          =   4290
            Left            =   -180
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   6840
            TabIndex        =   12
            Top             =   540
            Width           =   555
         End
         Begin VB.Image ImgRigaSx 
            Height          =   1050
            Left            =   720
            Top             =   4680
            Width           =   6540
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quota reale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   3420
            TabIndex        =   11
            Top             =   540
            Width           =   1425
         End
         Begin VB.Image Fil1_fine 
            Height          =   2490
            Left            =   6660
            Top             =   1800
            Width           =   1020
         End
         Begin VB.Label RealPos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-300,000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   4920
            TabIndex        =   10
            Top             =   450
            Width           =   1845
         End
         Begin VB.Image Fil1_corpo 
            Height          =   2505
            Left            =   3780
            Stretch         =   -1  'True
            Top             =   1860
            Width           =   3330
         End
         Begin VB.Image ImgFiletto1 
            Height          =   2520
            Left            =   2100
            Top             =   1920
            Width           =   4785
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   10
            Left            =   1500
            Top             =   450
            Width           =   3465
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7275
         Index           =   1
         Left            =   7740
         TabIndex        =   3
         Top             =   570
         Width           =   7635
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   7
            Left            =   7020
            TabIndex        =   62
            Top             =   3780
            Width           =   165
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   6
            Left            =   6840
            TabIndex        =   61
            Top             =   3780
            Width           =   165
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   5
            Left            =   6660
            TabIndex        =   60
            Top             =   3780
            Width           =   165
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   165
            Index           =   4
            Left            =   6480
            TabIndex        =   59
            Top             =   3780
            Width           =   165
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Index           =   3
            Left            =   6240
            TabIndex        =   52
            Top             =   3300
            Width           =   1215
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "RPM"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   6450
            TabIndex        =   51
            Top             =   3030
            Width           =   795
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Index           =   1
            Left            =   6240
            TabIndex        =   48
            Top             =   2670
            Width           =   1215
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "Passo"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   6420
            TabIndex        =   46
            Top             =   2430
            Width           =   795
         End
         Begin VB.Label Label9 
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "Quota start"
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
            Height          =   645
            Left            =   5190
            TabIndex        =   34
            Top             =   5940
            Width           =   3255
         End
         Begin VB.Label Label8 
            BackColor       =   &H00EE3959&
            BackStyle       =   0  'Transparent
            Caption         =   "Quota stop"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   585
            Left            =   2520
            TabIndex        =   33
            Top             =   5970
            Width           =   3255
         End
         Begin VB.Label Label01 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0      1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   360
            TabIndex        =   25
            Top             =   5730
            Width           =   1215
         End
         Begin VB.Image Selettore 
            Height          =   1155
            Index           =   1
            Left            =   360
            Top             =   6030
            Width           =   1185
         End
         Begin VB.Image ImgStop2 
            Height          =   4920
            Left            =   2640
            Top             =   930
            Width           =   1110
         End
         Begin VB.Image Imgstart2 
            Height          =   4515
            Left            =   5700
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Image ImgFilDx 
            Height          =   4260
            Left            =   5220
            Top             =   1020
            Width           =   2685
         End
         Begin VB.Image Fil2_base 
            Height          =   2565
            Left            =   3420
            Top             =   1800
            Width           =   780
         End
         Begin VB.Label StartQuoteDisplay 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   5580
            TabIndex        =   8
            Top             =   6480
            Width           =   1455
         End
         Begin VB.Label StopQuoteDisplay 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   7
            Top             =   6480
            Width           =   1455
         End
         Begin VB.Label RealPos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-300,000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   4950
            TabIndex        =   6
            Top             =   450
            Width           =   1845
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "mm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   6900
            TabIndex        =   5
            Top             =   510
            Width           =   555
         End
         Begin VB.Image Fil2_fine 
            Height          =   2475
            Left            =   0
            Top             =   1800
            Width           =   960
         End
         Begin VB.Image ImgRigaDx 
            Height          =   1140
            Left            =   600
            Top             =   4740
            Width           =   6180
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quota reale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   3420
            TabIndex        =   4
            Top             =   570
            Width           =   1425
         End
         Begin VB.Image Fil2_corpo 
            Height          =   2535
            Left            =   360
            Stretch         =   -1  'True
            Top             =   1740
            Width           =   3330
         End
         Begin VB.Image ImgFiletto2 
            Height          =   2430
            Left            =   840
            Top             =   1920
            Width           =   4755
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   0
            Left            =   1500
            Top             =   450
            Width           =   3465
         End
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Socket screwing"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   645
         Left            =   -69540
         TabIndex        =   54
         Top             =   2700
         Width           =   4575
      End
      Begin VB.Label NumGiriManicotto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68040
         TabIndex        =   53
         Top             =   3420
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Passo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   180
         TabIndex        =   44
         Top             =   1170
         Width           =   825
      End
      Begin VB.Label TempoVernDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -66390
         TabIndex        =   43
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo verniciatura (s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   645
         Left            =   -67980
         TabIndex        =   42
         Top             =   1020
         Width           =   5055
      End
      Begin VB.Label SpessManicottoDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -68010
         TabIndex        =   41
         Top             =   1590
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Spessore manicotto (mm)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   645
         Index           =   0
         Left            =   -69510
         TabIndex        =   40
         Top             =   900
         Width           =   4575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Verniciatrice 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Left            =   -71250
         TabIndex        =   39
         Top             =   1050
         Width           =   3375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Verniciatrice 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1395
         Left            =   -74850
         TabIndex        =   38
         Top             =   1050
         Width           =   3375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Manicottatrice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   -74430
         TabIndex        =   37
         Top             =   900
         Width           =   3795
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Tappatrice 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   -71730
         TabIndex        =   36
         Top             =   930
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Tappatrice 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   765
         Left            =   -74700
         TabIndex        =   35
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   -70350
         TabIndex        =   30
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   -73980
         TabIndex        =   29
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   6
         Left            =   -70350
         Top             =   2250
         Width           =   1185
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   5
         Left            =   -73980
         Top             =   2250
         Width           =   1185
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   -73260
         TabIndex        =   28
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   4
         Left            =   -73260
         Top             =   1950
         Width           =   1185
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -73710
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   -70740
         TabIndex        =   26
         Top             =   1710
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   3
         Left            =   -70710
         Top             =   2160
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   2
         Left            =   -73710
         Top             =   2130
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   80
      ImageHeight     =   80
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FilettoForm.frx":2224
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FilettoForm.frx":2827
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   75
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Label ThreadTypeDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   1860
      Width           =   3255
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Speed (mm/s)"
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
      Index           =   4
      Left            =   11970
      TabIndex        =   22
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lenght (mm)"
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
      Left            =   9330
      TabIndex        =   21
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step (mm)"
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
      Index           =   0
      Left            =   7800
      TabIndex        =   20
      Top             =   1500
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   4770
      TabIndex        =   19
      Top             =   1500
      Width           =   1455
   End
   Begin VB.Label ThreadStepDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   495
      Left            =   7680
      TabIndex        =   18
      Top             =   1860
      Width           =   1905
   End
   Begin VB.Label ThredLenghtDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   495
      Left            =   10050
      TabIndex        =   17
      Top             =   1860
      Width           =   1515
   End
   Begin VB.Label ThreadSpeedDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   12390
      TabIndex        =   16
      Top             =   1860
      Width           =   1515
   End
   Begin VB.Label LblData 
      BackStyle       =   0  'Transparent
      Caption         =   "......"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12270
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label LblOra 
      BackStyle       =   0  'Transparent
      Caption         =   "......."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   12750
      TabIndex        =   0
      Top             =   420
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   1170
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   13125
   End
End
Attribute VB_Name = "FilettoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Pos0start1 = 1470
Const posMin20_fil1 = 570
Const Pos0_fil1 = 1590
Const posMax50_fil1 = 3840

Const Pos0start2 = 4800
Const posMin20_fil2 = 5700
Const Pos0_fil2 = 4890
Const posMax50_fil2 = 2640

Const Kpos = (posMax50_fil1 - Pos0_fil1) / 50

Public FilettoType, FilettoStep, FilettoLenght, FilettoSpeed, QuotaStop, FilettoN_giri_man
Public CancBook As Boolean

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
    frmKernel.PaginaCorrente = Index
End Sub

Private Sub NumGiriManicotto_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 5000
    TOUCHNumericPad.Dati = DB460.Word(26)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        NumGiriManicotto.Caption = TOUCHNumericPad.Dati
        DB460.Word(26) = NumGiriManicotto.Caption
    End If
End Sub

Private Sub ThreadSpeedDisplay_Click()
    TOUCHNumericPad.Decimali = 0
     TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = ThreadSpeedDisplay.Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        ThreadSpeedDisplay.Caption = TOUCHNumericPad.Dati
        DB460.Word(66) = ThreadSpeedDisplay.Caption
    End If
End Sub

Private Sub TimerLocale_Timer()
        
             ' aggiorna i dati pagina
        lblbar(2) = PaginaWb.Ordine_Descrizione
        lblbar(4) = PaginaWb.Ricetta_Descrizione
        ' aggiorna lo stato del pulsante comunicazione
        If frmKernel.StatoCom Then
           If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
        Else
            If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
        End If
               
        If DB460.DatiCambiati Then
           Me.Update
           DB460.DatiCambiati = False
        End If
               
       ' lettura filetto reale delle 2 teste
       
       Label20(0).Caption = DB416.DWord(42) / 10
       Label20(1).Caption = DB417.DWord(42) / 10
       Label20(2).Caption = DB416.Word(28)
       Label20(3).Caption = DB417.Word(28)
       RealPos(0).Caption = Format(Val(Left(Str(DB416.DWord(34)), 8)) / 1000, "#000.000")
       RealPos(1).Caption = Format(Val(Left(Str(DB417.DWord(34)), 8)) / 1000, "#000.000")
      ' NumGiriManicotto.Caption = DB460.Word(26)
       
       Select Case DB416.Word(46)
       Case 0
          Label24(0).BackColor = &H80000005
          Label24(1).BackColor = &H80000005
          Label24(2).BackColor = &H80000005
          Label24(3).BackColor = &H80000005
       Case 1
          Label24(0).BackColor = &H80FF80
          Label24(1).BackColor = &H80000005
          Label24(2).BackColor = &H80000005
          Label24(3).BackColor = &H80000005
       Case 2
        Label24(0).BackColor = &H80FF80
          Label24(1).BackColor = &H80FF80
          Label24(2).BackColor = &H80000005
          Label24(3).BackColor = &H80000005
       Case 3
        Label24(0).BackColor = &H80FF80
          Label24(1).BackColor = &H80FF80
          Label24(2).BackColor = &H80FF80
          Label24(3).BackColor = &H80000005
       Case 4
           Label24(0).BackColor = &H80FF80
          Label24(1).BackColor = &H80FF80
          Label24(2).BackColor = &H80FF80
          Label24(3).BackColor = &H80FF80
       End Select
       
       Select Case DB417.Word(46)
       Case 0
          Label24(4).BackColor = &H80000005
          Label24(5).BackColor = &H80000005
          Label24(6).BackColor = &H80000005
          Label24(7).BackColor = &H80000005
       Case 1
          Label24(7).BackColor = &H80FF80
          Label24(6).BackColor = &H80000005
          Label24(5).BackColor = &H80000005
          Label24(4).BackColor = &H80000005
       Case 2
          Label24(7).BackColor = &H80FF80
          Label24(6).BackColor = &H80FF80
          Label24(5).BackColor = &H80000005
          Label24(4).BackColor = &H80000005
       Case 3
          Label24(7).BackColor = &H80FF80
          Label24(6).BackColor = &H80FF80
          Label24(5).BackColor = &H80FF80
          Label24(4).BackColor = &H80000005
       Case 4
           Label24(7).BackColor = &H80FF80
          Label24(6).BackColor = &H80FF80
          Label24(5).BackColor = &H80FF80
          Label24(4).BackColor = &H80FF80
       End Select
End Sub

'FUNZIONE DI AGGIORNAMENTO DEI CONTROLLI CONTENUTI NELLA PAGINA

Sub Update()
      Dim i
    
       For i = 0 To 1
            StartQuoteDisplay(i).Caption = DB460.DWord(48 + i * 40) / 1000
            StopQuoteDisplay(i).Caption = DB460.DWord(52 + i * 40) / 1000
            ThreadStepDisplay.Caption = DB460.DWord(56) / 10000
            ThreadSpeedDisplay.Caption = DB460.Word(66)
            
            If DB460.Bit(78 + i * 40, 0) = True Then
                Selettore(i).Picture = ImageList1.ListImages(2).Picture
            Else
                Selettore(i).Picture = ImageList1.ListImages(1).Picture
            End If
      Next
      
      SpessManicottoDisplay(0).Caption = DB460.Word(24)
      TempoVernDisplay(0).Caption = DB460.Word(32)
      NumGiriManicotto.Caption = DB460.Word(26)
      
       ' filettatrici 1 e 2
       If DB460.Bit(78, 0) = False Then
         Frame2(0).BackColor = &HC0C0C0
       Else
         Frame2(0).BackColor = &H80FFFF
       End If
    
       If DB460.Bit(118, 0) = False Then
         Frame2(1).BackColor = &HC0C0C0
       Else
         Frame2(1).BackColor = &H80FFFF
       End If
       'Tappatrice 1
       If DB460.Bit(28, 2) = True Then
          Selettore(2).Picture = ImageList1.ListImages(2).Picture
       Else
          Selettore(2).Picture = ImageList1.ListImages(1).Picture
       End If
       'Tappatrice 2
       If DB460.Bit(28, 3) = True Then
          Selettore(3).Picture = ImageList1.ListImages(2).Picture
       Else
          Selettore(3).Picture = ImageList1.ListImages(1).Picture
       End If
       ' manicottatrice
       If DB460.Bit(28, 4) = True Then
          Selettore(4).Picture = ImageList1.ListImages(2).Picture
       Else
          Selettore(4).Picture = ImageList1.ListImages(1).Picture
       End If
        ' verniciatrice 1
       If DB460.Bit(28, 0) = True Then
          Selettore(5).Picture = ImageList1.ListImages(2).Picture
       Else
          Selettore(5).Picture = ImageList1.ListImages(1).Picture
       End If
        ' verniciatrice 2
       If DB460.Bit(28, 1) = True Then
          Selettore(6).Picture = ImageList1.ListImages(2).Picture
       Else
          Selettore(6).Picture = ImageList1.ListImages(1).Picture
       End If
End Sub

Private Sub Form_Load()
        
   LblData.Top = 240
   LblOra.Top = 480
   LblOra.Left = 12480
   LblData.Left = 12000
      
   'Picture = LoadPicture("..\bitmap\SfondoDP6_ver2_0.jpg")
   ImgFilSx.Picture = LoadPicture("..\bitmap\Filettatura\filettatrice1.gif")
   ImgFilDx.Picture = LoadPicture("..\bitmap\Filettatura\filettatrice2.gif")
   ImgFiletto1.Picture = LoadPicture("..\bitmap\Filettatura\filetto1.gif")
   ImgFiletto2.Picture = LoadPicture("..\bitmap\Filettatura\filetto2.gif")
   Fil1_fine.Picture = LoadPicture("..\bitmap\Filettatura\tubofine1.gif")
   Fil2_fine.Picture = LoadPicture("..\bitmap\Filettatura\tubofine2.gif")
   Fil1_base.Picture = LoadPicture("..\bitmap\Filettatura\tuboinizio1.gif")
   Fil2_base.Picture = LoadPicture("..\bitmap\Filettatura\tuboinizio2.gif")
   Fil1_corpo.Picture = LoadPicture("..\bitmap\Filettatura\tubo1.gif")
   Fil2_corpo.Picture = LoadPicture("..\bitmap\Filettatura\tubo2.gif")
   ImgStart1.Picture = LoadPicture("..\bitmap\Filettatura\inizio.gif")
   Imgstart2.Picture = LoadPicture("..\bitmap\Filettatura\inizio.gif")
   ImgStop1.Picture = LoadPicture("..\bitmap\Filettatura\fine.gif")
   ImgStop2.Picture = LoadPicture("..\bitmap\Filettatura\fine.gif")
   ImgRigaSx.Picture = LoadPicture("..\bitmap\Filettatura\righello1.gif")
   ImgRigaDx.Picture = LoadPicture("..\bitmap\Filettatura\righello2.gif")
   
   WindowState = 2
   
    ImgStop1.Move Pos0_fil1 + StopQuoteDisplay(1) * Kpos
    ImgStart1.Move Pos0start1 - Abs(StartQuoteDisplay(1) * Kpos)
    Fil1_base.Move ImgStop1.Left + 400
    Fil1_corpo.Move Fil1_base.Left + 200
    Fil1_corpo.Width = Fil1_fine.Left - Fil1_corpo.Left + 500
    
    ImgStop2.Move Pos0_fil2 - StopQuoteDisplay(0) * Kpos
    Imgstart2.Move Pos0start2 + Abs(StartQuoteDisplay(0) * Kpos)
    Fil2_base.Move ImgStop2.Left - 100
    Fil2_corpo.Move Fil2_fine.Left + Fil2_fine.Width - 400
    Fil2_corpo.Width = Fil2_base.Left - Fil2_fine.Left + Fil2_fine.Width - 700
    
    ' scritte multilingua
    
    ScritteMultilingua
End Sub
Private Sub Form_Activate()
    TimerLocale.Enabled = True
    TimerLocale.Interval = 500
   
    Barra21.Selezionato = 9
'============================================================
' CARICA IL FILETTO DAL DATABASE
'============================================================

  On Error GoTo MancaBookmark
      
   With AdoFiletti

        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Produzione.mdb;Persist Security Info=False"
        .CommandType = adCmdText
        .RecordSource = "SELECT Filetti.Type, Filetti.Step, Filetti.Lenght, Filetti.Speed,Filetti.Bookmark,Filetti.NGiriMan  FROM Filetti;"
        .Refresh
        
        .Recordset.Find ("Bookmark=1")
        ThreadTypeDisplay = .Recordset.Fields("Type")
        On Error GoTo 0
        
        Set .Recordset.ActiveConnection = Nothing
    End With
      ' abilitazione temporizzatore locale
    TimerLocale.Enabled = True
    ' (disattivare in Form_deactivate)
    Exit Sub
MancaBookmark:
     MsgBox "Manca Bookmark in tabella filetti"
End Sub

Private Sub Form_Deactivate()
TimerLocale.Enabled = False
End Sub

'Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
'   frmKernel.PaginaCorrente = Index
'End Sub

'============================================================
' CARICA LA PAGINA DI SCELTA TIPO FILETTO
'============================================================

Private Sub Command1_Click()
    TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
   TechPasswordForm.Show vbModal
   If TechPasswordForm.LoginSucceeded = False Then Exit Sub
   Unload TechPasswordForm
   DoEvents
   DialogFiletti.Show vbModal
End Sub

Private Sub Selettore_Click(Index As Integer)
           
      Select Case Index
      Case 0, 1
          DB460.Bit(78 + Index * 40, 0) = Not DB460.Bit(78 + Index * 40, 0)
          If DB460.Bit(78 + Index * 40, 0) = True Then
              Selettore(Index).Picture = ImageList1.ListImages(2).Picture
          Else
             Selettore(Index).Picture = ImageList1.ListImages(1).Picture
          End If
      Case 2
          DB460.Bit(28, 2) = Not DB460.Bit(28, 2)
          If DB460.Bit(28, 2) = True Then
              Selettore(Index).Picture = ImageList1.ListImages(2).Picture
          Else
              Selettore(Index).Picture = ImageList1.ListImages(1).Picture
          End If
      Case 3
          DB460.Bit(28, 3) = Not DB460.Bit(28, 3)
          If DB460.Bit(28, 3) = True Then
              Selettore(Index).Picture = ImageList1.ListImages(2).Picture
          Else
              Selettore(Index).Picture = ImageList1.ListImages(1).Picture
          End If
       Case 4
          DB460.Bit(28, 4) = Not DB460.Bit(28, 4)
          If DB460.Bit(28, 4) = True Then
              Selettore(Index).Picture = ImageList1.ListImages(2).Picture
          Else
              Selettore(Index).Picture = ImageList1.ListImages(1).Picture
          End If
      Case 5
          DB460.Bit(28, 0) = Not DB460.Bit(28, 0)
          If DB460.Bit(28, 0) = True Then
              Selettore(Index).Picture = ImageList1.ListImages(2).Picture
          Else
              Selettore(Index).Picture = ImageList1.ListImages(1).Picture
          End If
      Case 6
          DB460.Bit(28, 1) = Not DB460.Bit(28, 1)
          If DB460.Bit(28, 1) = True Then
              Selettore(Index).Picture = ImageList1.ListImages(2).Picture
          Else
              Selettore(Index).Picture = ImageList1.ListImages(1).Picture
          End If
      End Select
End Sub

Private Sub SpessManicottoDisplay_Click(Index As Integer)
    TOUCHNumericPad.Decimali = 0
     TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 10
    TOUCHNumericPad.Dati = SpessManicottoDisplay(0).Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        SpessManicottoDisplay(0).Caption = TOUCHNumericPad.Dati
        DB460.Word(24) = SpessManicottoDisplay(0).Caption
    End If
End Sub

Private Sub StartQuoteDisplay_Click(Index As Integer)
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = -20
    TOUCHNumericPad.ValoreMax = 0
    TOUCHNumericPad.Dati = StartQuoteDisplay(Index).Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        StartQuoteDisplay(Index).Caption = -Abs(TOUCHNumericPad.Dati)
        DB460.DWord(48 + Index * 40) = Abs(StartQuoteDisplay(Index).Caption * 1000)
    End If
End Sub
Private Sub StartQuoteDisplay_Change(Index As Integer)
      
      ImgStop1.Move Pos0_fil1 + StopQuoteDisplay(1) * Kpos
      ImgStart1.Move Pos0start1 - Abs(StartQuoteDisplay(1) * Kpos)
      Fil1_base.Move ImgStop1.Left + 400
      Fil1_corpo.Move Fil1_base.Left + 200
      Fil1_corpo.Width = Fil1_fine.Left - Fil1_corpo.Left + 500
      
      ImgStop2.Move Pos0_fil2 - StopQuoteDisplay(0) * Kpos
      Imgstart2.Move Pos0start2 + Abs(StartQuoteDisplay(0) * Kpos)
      Fil2_base.Move ImgStop2.Left - 100
      Fil2_corpo.Move Fil2_fine.Left + Fil2_fine.Width - 400
      Fil2_corpo.Width = Fil2_base.Left - Fil2_fine.Left + Fil2_fine.Width - 700
      
End Sub
Private Sub StopQuoteDisplay_Change(Index As Integer)
      
      ImgStop1.Move Pos0_fil1 + StopQuoteDisplay(1) * Kpos
      ImgStart1.Move Pos0start1 - Abs(StartQuoteDisplay(1) * Kpos)
      Fil1_base.Move ImgStop1.Left + 400
      Fil1_corpo.Move Fil1_base.Left + 200
      Fil1_corpo.Width = Fil1_fine.Left - Fil1_corpo.Left + 500
      
      ImgStop2.Move Pos0_fil2 - StopQuoteDisplay(0) * Kpos
      Imgstart2.Move Pos0start2 + Abs(StartQuoteDisplay(0) * Kpos)
      Fil2_base.Move ImgStop2.Left - 100
      Fil2_corpo.Move Fil2_fine.Left + Fil2_fine.Width - 400
      Fil2_corpo.Width = Fil2_base.Left - Fil2_fine.Left + Fil2_fine.Width - 700

End Sub

Private Sub StopQuoteDisplay_Click(Index As Integer)
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 50
    TOUCHNumericPad.Dati = StopQuoteDisplay(Index).Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        StopQuoteDisplay(Index).Caption = TOUCHNumericPad.Dati
        DB460.DWord(52 + Index * 40) = StopQuoteDisplay(Index).Caption * 1000
    End If
End Sub

Sub Download()
Dim i

   For i = 0 To 1
      FilettoForm.StartQuoteDisplay(i).Caption = 0
      FilettoForm.StopQuoteDisplay(i).Caption = QuotaStop
      DB460.DWord(48 + i * 40) = FilettoForm.StartQuoteDisplay(i).Caption * 1000
      DB460.DWord(52 + i * 40) = FilettoForm.StopQuoteDisplay(i).Caption * 1000
   Next
   
   ThreadTypeDisplay.Caption = FilettoType
   ThreadStepDisplay.Caption = FilettoStep
   ThredLenghtDisplay.Caption = FilettoLenght
   ThreadSpeedDisplay.Caption = FilettoSpeed
   NumGiriManicotto.Caption = FilettoN_giri_man
   
   DB460.DWord(56) = FilettoStep * 10000
   DB460.Word(66) = FilettoSpeed
   DB460.Word(26) = FilettoN_giri_man
End Sub

Private Sub AdoFiletti_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
   
   If CancBook = False Then
         With AdoFiletti
              'On Error GoTo Error
              FilettoType = .Recordset("Type")
              FilettoStep = .Recordset("Step")
              FilettoLenght = .Recordset("Lenght")
              FilettoSpeed = .Recordset("Speed")
              FilettoN_giri_man = .Recordset("NGiriMan")
              QuotaStop = .Recordset("Qstop")
              ' bit per il PLC di conferma cambio filetto
              DB460.Bit(28, 4) = True    'presente nella 2202
              DB460.Bit(78, 1) = True
              DB460.Bit(118, 1) = True
        End With
          
        Call Download
        End If
        Exit Sub
Error:
      MsgBox "Errore nel DB filetti"
      On Error GoTo 0
End Sub

Private Sub TempoVernDisplay_Click(Index As Integer)
    TOUCHNumericPad.Decimali = 0
     TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 10
    TOUCHNumericPad.Dati = TempoVernDisplay(0).Caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        TempoVernDisplay(0).Caption = TOUCHNumericPad.Dati
        DB460.Word(32) = TempoVernDisplay(0).Caption
    End If
End Sub

Sub ScritteMultilingua()
    Label4.Caption = Param.Text("Type")
    Label3(0).Caption = Param.Text("Step (mm)")
    Label2.Caption = Param.Text("Lenght (mm)")
    Label20(4).Caption = Param.Text("Speed (mm/s)")
    SSTab1.TabCaption(0) = Param.Text("Filettatrici")
    SSTab1.TabCaption(1) = Param.Text("Tappatrici")
    SSTab1.TabCaption(2) = Param.Text("Manicottatrici")
    SSTab1.TabCaption(3) = Param.Text("Verniciatrici")
    Label3(1).Caption = Param.Text("Quota reale")
    Label3(2).Caption = Param.Text("Quota reale")
    Label6.Caption = Param.Text("Quota start")
    Label9.Caption = Label6.Caption
    Label7.Caption = Param.Text("Quota stop")
    Label8.Caption = Label7.Caption
    Label10.Caption = Param.Text("Tappatrice") & " 1"
    Label11.Caption = Param.Text("Tappatrice") & " 2"
    Label12.Caption = Param.Text("Manicottatrice")
    Label18.Caption = Param.Text("Passo")
    Label19.Caption = Param.Text("Passo")
    Label15(0) = Param.Text("Spessore manicotto (mm)")
    Label13.Caption = Param.Text("Verniciatrice") & " 1"
    Label14.Caption = Param.Text("Verniciatrice") & " 2"
    Label16.Caption = Param.Text("Tempo verniciatura (s)")
    Command1.Caption = Param.Text("Cambio filetto")
    lblbar(5) = Param.Text("Thred page")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
    lblbar(1) = Param.Text("Pagina")
End Sub

Private Sub XPButton1_Click(Index As Integer)
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
           Set frmHelp = Nothing
           frmHelp.Errori = False
           frmHelp.NomeFile = "TIP.HTM"
           frmHelp.Contesto = "DP6 : CP_L2_1 COM LOG"
           frmHelp.Top = 0
           frmHelp.Left = 7500
           frmHelp.Show vbModal
     Case 3
  On Error GoTo ErrorePercorso
           Unload frmHelp
           Set frmHelp = Nothing
           With frmHelp
                .Errori = True
                .NomeFile = "Filetto_pagina.htm"
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .webPreview.Move 0, 0, frmHelp.Width, frmHelp.Height
               .Show
           End With
  End Select
  Exit Sub
ErrorePercorso:
           MsgBox "Percorso file guida errato", vbExclamation
End Sub
