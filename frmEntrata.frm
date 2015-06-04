VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmEntrata 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "100"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9315
      Left            =   90
      TabIndex        =   16
      Top             =   1110
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   16431
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   1058
      BackColor       =   13160660
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Produzione"
      TabPicture(0)   =   "frmEntrata.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "MSChartTubiMin"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Impostazioni"
      TabPicture(1)   =   "frmEntrata.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Parametri"
      TabPicture(2)   =   "frmEntrata.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameChains"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tubo test"
      TabPicture(3)   =   "frmEntrata.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin VB.Frame FrameChains 
         Height          =   7635
         Left            =   -70410
         TabIndex        =   53
         Top             =   1200
         Width           =   6255
         Begin VB.CommandButton cmdTubeSquare 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Square"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   1740
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   5220
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.CommandButton cmdTubeRound 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Round"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   750
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   5220
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.ComboBox ComboChain 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            ItemData        =   "frmEntrata.frx":0070
            Left            =   1110
            List            =   "frmEntrata.frx":0072
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1290
            Width           =   3855
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Target"
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
            Index           =   21
            Left            =   3660
            TabIndex        =   69
            Top             =   3300
            Width           =   765
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Real"
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
            Index           =   20
            Left            =   2010
            TabIndex        =   68
            Top             =   3300
            Width           =   525
         End
         Begin VB.Label lblRollWay_real 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1440
            TabIndex        =   67
            Top             =   3600
            Width           =   1665
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tube height [in]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   19
            Left            =   3270
            TabIndex        =   66
            Top             =   6120
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lbl_tubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3360
            TabIndex        =   65
            Top             =   6960
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tube shape"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   18
            Left            =   360
            TabIndex        =   62
            Top             =   4470
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tube width [in]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   17
            Left            =   360
            TabIndex        =   61
            Top             =   6120
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lbl_tubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   450
            TabIndex        =   60
            Top             =   6960
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tube length [in]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   16
            Left            =   3330
            TabIndex        =   59
            Top             =   4470
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label lbl_tubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3420
            TabIndex        =   58
            Top             =   5340
            Visible         =   0   'False
            Width           =   2475
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Rollway height position [in]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   15
            Left            =   1770
            TabIndex        =   57
            Top             =   2310
            Width           =   2655
         End
         Begin VB.Label LblRollwayPos 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3180
            TabIndex        =   56
            Top             =   3600
            Width           =   1665
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Chains select"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   10
            Left            =   1800
            TabIndex        =   55
            Top             =   390
            Width           =   2655
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3975
         Left            =   -74490
         TabIndex        =   50
         Top             =   1200
         Width           =   3285
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entry tube sensor  [mm]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Index           =   14
            Left            =   420
            TabIndex        =   52
            Top             =   1050
            Width           =   2490
         End
         Begin VB.Label LblTastatore 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   420
            TabIndex        =   51
            Top             =   1740
            Width           =   2505
         End
      End
      Begin VB.Frame Frame4 
         Height          =   7725
         Left            =   -70410
         TabIndex        =   40
         Top             =   1110
         Visible         =   0   'False
         Width           =   6255
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080FFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   3210
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1770
            Width           =   1845
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000C000&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   1050
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1770
            Width           =   1845
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0080FFFF&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   3210
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H0000C000&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   1050
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1080
            Width           =   1845
         End
         Begin dp6.vbalProgressBar vbalProgressBar5 
            Height          =   3975
            Left            =   1800
            TabIndex        =   45
            Top             =   3630
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   7011
            Picture         =   "frmEntrata.frx":0074
            BackColor       =   12632256
            ForeColor       =   0
            Appearance      =   2
            BarPicture      =   "frmEntrata.frx":0090
            BarPictureMode  =   0
            BackPictureMode =   0
            Value           =   50
            ShowText        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            XpStyle         =   -1  'True
         End
         Begin dp6.vbalProgressBar vbalProgressBar6 
            Height          =   3975
            Left            =   3870
            TabIndex        =   46
            Top             =   3630
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   7011
            Picture         =   "frmEntrata.frx":0404
            BackColor       =   12632256
            ForeColor       =   0
            Appearance      =   2
            BarPicture      =   "frmEntrata.frx":0420
            BarPictureMode  =   0
            BackPictureMode =   0
            Value           =   50
            ShowText        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            XpStyle         =   -1  'True
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tubes in the entry stack"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   13
            Left            =   2010
            TabIndex        =   49
            Top             =   240
            Width           =   2250
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tubes last order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   12
            Left            =   1260
            TabIndex        =   48
            Top             =   2820
            Width           =   1485
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total tubes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   11
            Left            =   3360
            TabIndex        =   47
            Top             =   2820
            Width           =   1485
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5145
         Left            =   750
         TabIndex        =   23
         Top             =   2070
         Width           =   13665
         Begin dp6.vbalProgressBar vbalProgressBar1 
            Height          =   3915
            Left            =   10560
            TabIndex        =   24
            Top             =   1020
            Visible         =   0   'False
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   6906
            Picture         =   "frmEntrata.frx":06F0
            BackColor       =   -2147483637
            ForeColor       =   0
            Appearance      =   0
            BarPicture      =   "frmEntrata.frx":070C
            BackPictureMode =   0
            Value           =   50
            ShowText        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            XpStyle         =   -1  'True
         End
         Begin dp6.vbalProgressBar vbalProgressBar2 
            Height          =   3915
            Left            =   12150
            TabIndex        =   25
            Top             =   1020
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   6906
            Picture         =   "frmEntrata.frx":0A8B
            BackColor       =   -2147483637
            ForeColor       =   0
            Appearance      =   0
            BarPicture      =   "frmEntrata.frx":0AA7
            BackPictureMode =   0
            Value           =   50
            ShowText        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            XpStyle         =   -1  'True
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   1
            Visible         =   0   'False
            X1              =   2970
            X2              =   3570
            Y1              =   3540
            Y2              =   3030
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FF00&
            BorderWidth     =   2
            Index           =   0
            X1              =   3570
            X2              =   2970
            Y1              =   2970
            Y2              =   2490
         End
         Begin VB.Line Line5 
            X1              =   10140
            X2              =   13020
            Y1              =   4980
            Y2              =   4980
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vel. tubificio [m/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   7
            Left            =   10050
            TabIndex        =   39
            Top             =   180
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vel. via rulli [m/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   6
            Left            =   11700
            TabIndex        =   38
            Top             =   180
            Width           =   1545
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            BorderStyle     =   2  'Dash
            X1              =   3270
            X2              =   3270
            Y1              =   3330
            Y2              =   1110
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   3300
            X2              =   4170
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label LblRifVel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   4
            Left            =   7650
            TabIndex        =   37
            Top             =   2790
            Width           =   1665
         End
         Begin VB.Label LblRifVel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   4230
            TabIndex        =   36
            Top             =   2790
            Width           =   1665
         End
         Begin VB.Label LblRifVel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   2
            Left            =   450
            TabIndex        =   35
            Top             =   4590
            Width           =   1635
         End
         Begin VB.Label LblRifVel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   450
            TabIndex        =   34
            Top             =   3390
            Width           =   1635
         End
         Begin VB.Label LblRifVel 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   0
            Left            =   450
            TabIndex        =   33
            Top             =   2220
            Width           =   1635
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vel. via rulli [m/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   660
            Index           =   4
            Left            =   7650
            TabIndex        =   32
            Top             =   2130
            Width           =   1665
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Aumento [%]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   510
            Index           =   3
            Left            =   4230
            TabIndex        =   31
            Top             =   2280
            Width           =   1680
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sovravelocità [m/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   645
            Index           =   2
            Left            =   420
            TabIndex        =   30
            Top             =   3960
            Width           =   1680
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vel. manuale [m/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   630
            Index           =   1
            Left            =   420
            TabIndex        =   29
            Top             =   2760
            Width           =   1680
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vel. tubificio [m/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   0
            Left            =   420
            TabIndex        =   28
            Top             =   1560
            Width           =   1680
         End
         Begin VB.Label Label01Rif 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0        1"
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
            Left            =   4260
            TabIndex        =   27
            Top             =   840
            Width           =   1635
         End
         Begin VB.Image SelettoreRif 
            Height          =   1155
            Left            =   4470
            Top             =   1050
            Width           =   1185
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sel. Velocità in ingresso"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   720
            Index           =   5
            Left            =   4260
            TabIndex        =   26
            Top             =   180
            Width           =   1665
         End
         Begin VB.Image Image1 
            Height          =   2985
            Left            =   390
            Picture         =   "frmEntrata.frx":0E26
            Top             =   2130
            Width           =   9960
         End
         Begin VB.Shape ShapeRif 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   1575
            Left            =   4260
            Top             =   720
            Width           =   1665
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5355
         Left            =   -63300
         TabIndex        =   18
         Top             =   1920
         Width           =   3405
         Begin dp6.vbalProgressBar vbalProgressBar4 
            Height          =   4065
            Left            =   750
            TabIndex        =   19
            Top             =   1050
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   7170
            Picture         =   "frmEntrata.frx":13FE
            BackColor       =   12632256
            ForeColor       =   0
            Appearance      =   2
            BarPicture      =   "frmEntrata.frx":141A
            BarPictureMode  =   0
            BackPictureMode =   0
            Value           =   50
            ShowText        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            XpStyle         =   -1  'True
         End
         Begin dp6.vbalProgressBar vbalProgressBar3 
            Height          =   4035
            Left            =   2310
            TabIndex        =   20
            Top             =   1080
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   7117
            Picture         =   "frmEntrata.frx":1760
            BackColor       =   12632256
            ForeColor       =   0
            Appearance      =   2
            BarPicture      =   "frmEntrata.frx":177C
            BarPictureMode  =   0
            BackPictureMode =   0
            Value           =   50
            ShowText        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            XpStyle         =   -1  'True
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tubificio [Tubi/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   8
            Left            =   270
            TabIndex        =   22
            Top             =   300
            Width           =   1380
         End
         Begin VB.Label LblVel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Linea [Tubi/min]"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   675
            Index           =   9
            Left            =   1830
            TabIndex        =   21
            Top             =   300
            Width           =   1380
         End
      End
      Begin MSChart20Lib.MSChart MSChartTubiMin 
         Height          =   5265
         Left            =   -74760
         OleObjectBlob   =   "frmEntrata.frx":1AE0
         TabIndex        =   17
         Top             =   2010
         Width           =   11295
      End
   End
   Begin VB.TextBox TextValore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8940
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   10020
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TextOra 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   10020
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox Textdata 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6540
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   3
         Left            =   13800
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "frmEntrata.frx":3E94
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
         TabIndex        =   11
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   3150
      Top             =   6060
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2430
      Top             =   6030
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
            Picture         =   "frmEntrata.frx":5F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntrata.frx":6525
            Key             =   ""
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "frmEntrata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PassOK As Boolean
Private Chain As Integer

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub CmdPercorso_Click()
   frmAvvisi.AvvisoBypass = True
   frmAvvisi.Show vbModal
   TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
   TechPasswordForm.Show vbModal
   If TechPasswordForm.LoginSucceeded = False Then Exit Sub
   PassOK = True
   Unload TechPasswordForm
End Sub

Private Sub cmdTubeRound_Click()
    DB450.Bit(2, 0) = True
End Sub

Private Sub cmdTubeSquare_Click()
    DB450.Bit(2, 0) = False
End Sub

Private Sub ComboChain_Click()

    Select Case ComboChain.ListIndex
    Case 0
        DB450.Bit(22, 2) = 0
    DB450.Bit(22, 3) = 0
    DB450.Bit(22, 4) = 0
    DB450.Bit(22, 5) = 0
    DB450.Bit(22, 6) = 0
    Case 1
        DB450.Bit(22, 2) = 1
    DB450.Bit(22, 3) = 0
    DB450.Bit(22, 4) = 0
    DB450.Bit(22, 5) = 0
    DB450.Bit(22, 6) = 0
    Case 2
        DB450.Bit(22, 2) = 0
    DB450.Bit(22, 3) = 1
    DB450.Bit(22, 4) = 0
    DB450.Bit(22, 5) = 0
    DB450.Bit(22, 6) = 0
    Case 3
        DB450.Bit(22, 2) = 0
    DB450.Bit(22, 3) = 0
    DB450.Bit(22, 4) = 1
    DB450.Bit(22, 5) = 0
    DB450.Bit(22, 6) = 0
    Case 4
        DB450.Bit(22, 2) = 0
    DB450.Bit(22, 3) = 0
    DB450.Bit(22, 4) = 0
    DB450.Bit(22, 5) = 1
    DB450.Bit(22, 6) = 0
    Case 5
        DB450.Bit(22, 2) = 0
    DB450.Bit(22, 3) = 0
    DB450.Bit(22, 4) = 0
    DB450.Bit(22, 5) = 0
    DB450.Bit(22, 6) = 1
    Case Else
    
    End Select
End Sub

Private Sub Command1_Click(Index As Integer)
'    Select Case Index
'    Case 0
'        DB453.Bit(22, 4) = True
'    Case 2
'        DB453.Bit(22, 5) = True
'    Case 1
'        DB453.Bit(22, 6) = True
'    Case 3
'        DB453.Bit(22, 7) = True
'    End Select
End Sub

Private Sub lbl_tubeH_Click()
    TOUCHNumericPad.Decimali = 2
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par004_Tubo_AltezzaMin"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par003_Tubo_AltezzaMax"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.Dati = Conv_UM.Conversione(DB450.Word(8) / 10, UM.mm, UM.inch, 6)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
       DB450.Word(8) = Conv_UM.Conversione(TOUCHNumericPad.Dati * 10, UM.inch, UM.mm, 6)
    End If
End Sub

Private Sub lbl_tubeL_Click()
    TOUCHNumericPad.Decimali = 2
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par008_Tubo_LunghezzaMin"), UM.mt, UM.inch)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par007_Tubo_LunghezzaMax"), UM.mt, UM.inch)
    TOUCHNumericPad.Dati = Conv_UM.Conversione(DB450.Word(4), UM.mm, UM.ft, 6)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
       DB450.Word(4) = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.ft, UM.mm, 6)
    End If
End Sub

Private Sub lbl_tubeW_Click()
    TOUCHNumericPad.Decimali = 2
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par002_Tubo_LarghezzaMin"), UM.mt, UM.inch, 3) ' * 1000
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par001_Tubo_LarghezzaMax"), UM.mt, UM.inch, 3) '* 1000
    TOUCHNumericPad.Dati = Conv_UM.Conversione(DB450.Word(6) / 10, UM.mm, UM.inch, 6)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
       DB450.Word(6) = Conv_UM.Conversione(TOUCHNumericPad.Dati * 10, UM.inch, UM.mm, 6)
    End If
End Sub

Private Sub LblRifVel_Click(Index As Integer)
    TOUCHNumericPad.Decimali = 0
    Select Case Index
'    Case 0
'            TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(30, UM.m_min, UM.ft_min, 2)
'            TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(200, UM.m_min, UM.ft_min, 2)
'            TOUCHNumericPad.Dati = Conv_UM.Conversione(DB410.Word(28), UM.m_min, UM.ft_min, 2)
'            TOUCHNumericPad.Show vbModal
'            If TOUCHNumericPad.DatiConfermati Then
'               DB410.Word(28) = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.ft_min, UM.m_min, 0)
'            End If
    Case 1
            TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(30, UM.m_min, UM.ft_min, 2)
            TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(200, UM.m_min, UM.ft_min, 2)
            TOUCHNumericPad.Dati = Conv_UM.Conversione(DB450.Word(24), UM.m_min, UM.ft_min, 2)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
               DB450.Word(24) = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.ft_min, UM.m_min, 0)
            End If
    Case 2
            TOUCHNumericPad.ValoreMin = 0
            TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(200, UM.m_min, UM.ft_min, 2)
            TOUCHNumericPad.Dati = Conv_UM.Conversione(DB450.Word(26), UM.m_min, UM.ft_min, 2)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
               DB450.Word(26) = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.ft_min, UM.m_min, 2)
            End If
    Case 3
            TOUCHNumericPad.ValoreMin = 100
            TOUCHNumericPad.ValoreMax = 500
            TOUCHNumericPad.Dati = DB450.Word(28)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                DB450.Word(28) = TOUCHNumericPad.Dati
            End If
    End Select
End Sub

Private Sub LblRollwayPos_Click()
    TOUCHNumericPad.Decimali = 2
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(0, UM.mm, UM.inch)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(2000, UM.mm, UM.inch)
    TOUCHNumericPad.Dati = Conv_UM.Conversione(DB448.Word(4), UM.mm, UM.inch, 6)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
       DB448.Word(4) = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.inch, UM.mm, 6)
    End If
End Sub

Private Sub SelettoreRif_Click()
    If Not PassOK Then
        frmAvvisi.AvvisoBypass = True
        frmAvvisi.Show vbModal
        TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
        TechPasswordForm.Show vbModal
        If TechPasswordForm.LoginSucceeded = False Then Exit Sub
        PassOK = True
        Unload TechPasswordForm
        Param.SetBit "Par209_RifTubificio", Not (Param.GetBit("Par209_RifTubificio"))
        DB450.Bit(22, 0) = Param.GetBit("Par209_RifTubificio")
        If Param.GetBit("Par209_RifTubificio") Then
            SelettoreRif.Picture = ImageList1.ListImages(2).Picture
        Else
            SelettoreRif.Picture = ImageList1.ListImages(1).Picture
        End If
    Else
        Param.SetBit "Par209_RifTubificio", Not (Param.GetBit("Par209_RifTubificio"))
        DB450.Bit(22, 0) = Param.GetBit("Par209_RifTubificio")
        If Param.GetBit("Par209_RifTubificio") Then
            SelettoreRif.Picture = ImageList1.ListImages(2).Picture
        Else
            SelettoreRif.Picture = ImageList1.ListImages(1).Picture
        End If
    End If
End Sub

Private Sub TimerLocale_Timer()
   Me.Update
End Sub
Sub Update()
        
  '  Chain = Abs(DB450.Bit(22, 1)) + 2 * Abs(DB450.Bit(22, 2))
  '  ComboChain.ListIndex = Chain
  
   ' LblTastatore = DB413.Word(28)
    
  '  vbalProgressBar5.Max = 110
  '  vbalProgressBar5.Min = 0
  '  vbalProgressBar5.value = DB414.Word(68)
  '  vbalProgressBar5.Text = vbalProgressBar5.value
        
  '  vbalProgressBar6.Max = 110
  '  vbalProgressBar6.Min = 0
  '  vbalProgressBar6.value = DB414.Word(70)
  '  vbalProgressBar6.Text = vbalProgressBar6.value
    
    'aggiornamento dati della pagina
    
    LblRollwayPos = Conv_UM.Conversione(DB448.Word(4), UM.mm, UM.inch, 3)
    lblRollWay_real = Conv_UM.Conversione(DB410.Word(56), UM.mm, UM.inch, 3)
    lbl_tubeL = Conv_UM.Conversione(DB450.Word(4), UM.mm, UM.inch, 3)
    lbl_tubeW = Conv_UM.Conversione(DB450.Word(6) / 10, UM.mm, UM.inch, 2)
    lbl_tubeH = Conv_UM.Conversione(DB450.Word(8) / 10, UM.mm, UM.inch, 2)
    
    If DB450.Bit(2, 0) Then
       cmdTubeRound.BackColor = vbGreen
       cmdTubeSquare.BackColor = &HC0C0C0
    Else
       cmdTubeRound.BackColor = &HC0C0C0
       cmdTubeSquare.BackColor = vbGreen
    End If
    
    Textdata.Text = BufferTubiMin.Data
    
    vbalProgressBar3.Max = 100
    vbalProgressBar3.Min = 0
    vbalProgressBar3.value = BufferTubiMin.BufferData(1, 1)
    vbalProgressBar3.Text = BufferTubiMin.BufferData(1, 1)
    
    vbalProgressBar4.Max = 100
    vbalProgressBar4.Min = 0
    On Error Resume Next
    Dim a
    a = Int(Format(Val(LblRifVel(0)) / (DB450.Word(4) / 1000), "##0"))
    vbalProgressBar4.value = Int(Val(LblRifVel(0)) / (DB450.Word(4) / 1000))
    vbalProgressBar4.Text = a
    On Error GoTo 0
    TextOra.Text = Format(Time$, "hh.mm.ss")
    TextValore = BufferTubiMin.BufferData(1, 1)
    ' aggiorna i dati pagina
    lblbar(2) = PaginaEntrata.Ordine_Descrizione
    lblbar(4) = PaginaEntrata.Ricetta_Descrizione
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    
'    If PassOK Then
'        Param.SetBit "Par209_RifTubificio", Not (Param.GetBit("Par209_RifTubificio"))
'        DB450.Bit(22, 0) = Not DB450.Bit(22, 0)
'
'        If Param.GetBit("Par209_RifTubificio") Then
'            SelettoreRif.Picture = ImageList1.ListImages(2).Picture
'        Else
'            SelettoreRif.Picture = ImageList1.ListImages(1).Picture
'        End If
'       Label01Rif.Visible = True
'       SelettoreRif.Visible = True
'       ShapeRif.Visible = True
'       XPButton1(4).Visible = False
'    End If
     
     If Param.GetBit("Par209_RifTubificio") Then
         Line1(0).Visible = False
         Line1(1).Visible = True
         SelettoreRif.Picture = ImageList1.ListImages(2).Picture
     Else
         SelettoreRif.Picture = ImageList1.ListImages(1).Picture
         Line1(0).Visible = True
         Line1(1).Visible = False
     End If
     
     Param.SetBit "Par209_RifTubificio", DB450.Bit(22, 0)
     
'     If False Or DB450.DatiCambiati Or DB410.DatiCambiati Then
        LblRifVel(0) = Conv_UM.Conversione(DB410.Word(28), UM.m_min, UM.ft_min, 0)
        vbalProgressBar1.Max = 400
        vbalProgressBar1.Min = 0
        vbalProgressBar1.value = Val(LblRifVel(0))
        vbalProgressBar1.Text = Val(LblRifVel(0))
        LblRifVel(4) = Conv_UM.Conversione(DB410.Word(30), UM.m_min, UM.ft_min, 0)
        vbalProgressBar2.Max = 400
        vbalProgressBar2.Min = 0
        vbalProgressBar2.value = Val(LblRifVel(4))
        vbalProgressBar2.Text = Val(LblRifVel(4))
        LblRifVel(1) = Conv_UM.Conversione(DB450.Word(24), UM.m_min, UM.ft_min, 0)
        LblRifVel(2) = Conv_UM.Conversione(DB450.Word(26), UM.m_min, UM.ft_min, 0)
        LblRifVel(3) = DB450.Word(28) ' Conv_UM.Conversione(DB450.Word(28), UM.m_min, UM.ft_min, 2)
'        If ControlloOverrideVR1.Occupato = False Then
'           ControlloOverrideVR1.Value = DB450.Word(24)
'           ControlloOverrideVR1.Refresh
'        End If
'
'        If ControlloOverrideVR2.Occupato = False Then
'           ControlloOverrideVR2.Value = DB450.Word(26)
'           ControlloOverrideVR2.Refresh
'        End If
'
'        If ControlloMonobeam1.Occupato = False Then
'           ControlloMonobeam1.Value = DB450.Word(28)
'           ControlloMonobeam1.Refresh
'        End If
'
'        If ControlloMonobeam2.Occupato = False Then
'           ControlloMonobeam2.Value = DB450.Word(30)
'           ControlloMonobeam2.Refresh
'        End If
                
        
    
'        If Param.GetBit("Par212_AttivaGestioneBypass") = False Then Exit Sub
'
'        If DB450.Bit(22, 0) = True Then
'            SelettoreBypass.Picture = ImageList1.ListImages(2).Picture
'        Else
'            SelettoreBypass.Picture = ImageList1.ListImages(1).Picture
'        End If
'
'        If DB450.Bit(22, 0) = False Then
'           Command1(0).BackColor = &H8000000F
'        Else
'           Command1(0).BackColor = &HFF00&
'        End If
'        If DB450.Bit(22, 1) = False Then
'           Command1(1).BackColor = &H8000000F
'        Else
'           Command1(1).BackColor = &HFF00&
'        End If
'        If DB450.Bit(22, 2) = False Then
'           Command1(2).BackColor = &H8000000F
'        Else
'           Command1(2).BackColor = &HFF00&
'        End If
'        If DB450.Bit(22, 3) = False Then
'           Command1(3).BackColor = &H8000000F
'        Else
'           Command1(3).BackColor = &HFF00&
'        End If
        DB450.DatiCambiati = False
'    End If
End Sub

'Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
'   frmKernel.PaginaCorrente = Index
'End Sub


Private Sub Form_Activate()
    Static primo As Boolean
    
    On Error Resume Next
    If primo = False Then SSTab1.Tab = 2
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = False
    FrameChains.Visible = Param.GetBit("Par105_SelezioneCatene")
    If FrameChains.Visible Then
       ComboChain.Clear
       ComboChain.AddItem "Auto"
       ComboChain.AddItem "From mill"
       ComboChain.AddItem "3 " & Param.Text("000000083")
       ComboChain.AddItem "4 " & Param.Text("000000083")
       ComboChain.AddItem "5 " & Param.Text("000000083")
       ComboChain.AddItem "6 " & Param.Text("000000083")
       Chain = Abs(DB450.Bit(22, 2)) + 2 * Abs(DB450.Bit(22, 3)) + 3 * Abs(DB450.Bit(22, 4)) + 4 * Abs(DB450.Bit(22, 5)) + 5 * Abs(DB450.Bit(22, 6))
       ComboChain.ListIndex = Chain
    End If
    
   TimerLocale.Enabled = True
   TimerLocale.Interval = 500
   primo = True
   Barra21.Selezionato = 12
    ' abilitazione temporizzatore locale
   Call Update
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
    PassOK = False
End Sub

Private Sub Form_Load()
  '  Picture = LoadPicture("..\bitmap\SfondoDP6_ver2_0.jpg")
    Image1.Picture = LoadPicture("..\bitmap\VelEntrata.gif")
    ScritteMultilingua
End Sub

'selezione bypass

Private Sub SelettoreBypass_Click()
'    DB413.Bit(26, 7) = Not DB413.Bit(26, 7)
'
'    If DB413.Bit(26, 7) = True Then
'        SelettoreBypass.Picture = ImageList1.ListImages(2).Picture
'    Else
'        SelettoreBypass.Picture = ImageList1.ListImages(1).Picture
'    End If
End Sub

Public Sub AggiornaDaUpDownControl()
'    DB450.Word(24) = ControlloOverrideVR1.Value
'    DB450.Word(26) = ControlloOverrideVR2.Value
'    DB450.Word(28) = ControlloMonobeam1.Value
'    DB450.Word(30) = ControlloMonobeam2.Value
End Sub

Sub ScritteMultilingua()
  '  Label7.Caption = Param.Text("Velocità via rulli (s)") & " 1 (%)"
  '  Label6.Caption = Param.Text("Velocità via rulli (s)") & " 2 (%)"
  '  Label8.Caption = Param.Text("VelMonobeam") & " 1 (%)"
  '  Label9.Caption = Param.Text("VelMonobeam") & " 2 (%)"
  '  Label2.Caption = Param.Text("Bypass")
  '  CmdPercorso.Caption = Param.Text("Percorso")
  '  Command1(0).Caption = Param.Text("Caricatore")
  '  Command1(1).Caption = Param.Text("Verniciatrice")
  '  Label1.Caption = Param.Text("Entrata")
    lblbar(5) = Param.Text("Entry page")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
    lblbar(1) = Param.Text("Pagina")
    LblVel(0) = Param.Text("VelTubi") & IIf(Conv_UM.SI_metrico = False, " [ft/min]", " [m/min]")
    LblVel(1) = Param.Text("Velman") & IIf(Conv_UM.SI_metrico = False, " [ft/min]", " [m/min]")
    LblVel(2) = Param.Text("Svel") & IIf(Conv_UM.SI_metrico = False, " [ft/min]", " [m/min]")
    LblVel(3) = Param.Text("Aum")
    LblVel(4) = Param.Text("VVR") & IIf(Conv_UM.SI_metrico = False, " [ft/min]", " [m/min]")
    LblVel(5) = Param.Text("VelIN")
    LblVel(6) = Param.Text("VVR") & IIf(Conv_UM.SI_metrico = False, " [ft/min]", " [m/min]")
    LblVel(7) = Param.Text("VelTubi") & IIf(Conv_UM.SI_metrico = False, " [ft/min]", " [m/min]")
    LblVel(8) = Param.Text("Vtubif")
    LblVel(9) = Param.Text("VelLinea")
    LblVel(10) = Param.Text("000000084")
    SSTab1.TabCaption(0) = Param.Text("000000001")
    SSTab1.TabCaption(1) = Param.Text("000000002")
    SSTab1.TabCaption(2) = Param.Text("000000003")
    SSTab1.TabCaption(3) = Param.Text("000000004")
End Sub

Private Sub VelocWBDisplay_Click()
'   Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = 0
'    TOUCHNumericPad.ValoreMax = 999999
'    TOUCHNumericPad.Dati = DB413.DWord(28)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        DB413.DWord(28) = TOUCHNumericPad.Dati
'    End If
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
           frmKernel.PaginaCorrente = PagService
           frmConnecting.Hide
     Case 2
           On Error GoTo ErrorePercorso
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
                .NomeFile = "entrata_pagina.htm"
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .webPreview.Move 0, 0, frmHelp.Width, frmHelp.Height
               .Show
           End With
'     Case 4
'            frmAvvisi.AvvisoBypass = True
'            frmAvvisi.Show vbModal
'            TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
'            TechPasswordForm.Show vbModal
'            If TechPasswordForm.LoginSucceeded = False Then Exit Sub
'            PassOK = True
'            Unload TechPasswordForm
  End Select
  Exit Sub
ErrorePercorso:
           MsgBox "Percorso file guida errato", vbExclamation
End Sub
