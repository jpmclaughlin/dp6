VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form CommandForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "422"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   360
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   422
      Left            =   13950
      TabIndex        =   22
      Top             =   3450
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   2355
   End
   Begin VB.ComboBox ComboDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   3
      Left            =   300
      Style           =   2  'Dropdown List
      TabIndex        =   108
      Top             =   5370
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox ComboDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   2
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   107
      Top             =   4710
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox ComboDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   106
      Top             =   4020
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox ComboDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   105
      Top             =   3390
      Visible         =   0   'False
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1125
      Left            =   7500
      TabIndex        =   101
      Top             =   3510
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   1984
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   882
      BackColor       =   12632256
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Next bundles destinations"
      TabPicture(0)   =   "CommandForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ComboStorage"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.ComboBox ComboStorage 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "CommandForm.frx":001C
         Left            =   750
         List            =   "CommandForm.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   510
         Width           =   2805
      End
   End
   Begin VB.CheckBox AdjRolls 
      BackColor       =   &H008080FF&
      Caption         =   "   Adjust vertical rolls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   2
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   5010
      Width           =   1545
   End
   Begin VB.CheckBox AdjRolls 
      BackColor       =   &H008080FF&
      Caption         =   "   Adjust vertical rolls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   1
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   5010
      Width           =   1545
   End
   Begin VB.CheckBox AdjRolls 
      BackColor       =   &H008080FF&
      Caption         =   "   Adjust vertical rolls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   0
      Left            =   3090
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   5010
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   5730
      TabIndex        =   87
      Top             =   8220
      Width           =   3255
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1155
      Left            =   0
      TabIndex        =   18
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   1
         Left            =   1650
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   17
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
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   13
         Top             =   300
         Width           =   1185
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
         Index           =   3
         Left            =   7260
         TabIndex        =   12
         Top             =   660
         Width           =   3495
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
         Left            =   5760
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.Line Line2 
         X1              =   10860
         X2              =   10860
         Y1              =   180
         Y2              =   1020
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2003"
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
         Left            =   9600
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblOra 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   5760
         X2              =   10860
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Anno"
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
         Index           =   4
         Left            =   8580
         TabIndex        =   8
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2309"
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
         Height          =   465
         Index           =   2
         Left            =   7260
         TabIndex        =   7
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Impianto"
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
         Left            =   5760
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image Image3 
         Height          =   1050
         Left            =   3150
         Picture         =   "CommandForm.frx":0020
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   2
         Left            =   3180
         TabIndex        =   0
         Top             =   150
         Width           =   8985
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Impianto      2333"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   555
         Left            =   6060
         TabIndex        =   5
         Top             =   150
         Width           =   3825
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Anno                                 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   6390
         TabIndex        =   4
         Top             =   690
         Width           =   3375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Data ora"
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
         Height          =   375
         Index           =   0
         Left            =   10590
         TabIndex        =   3
         Top             =   240
         Width           =   3435
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "DP6.1"
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
         Height          =   375
         Index           =   1
         Left            =   10740
         TabIndex        =   2
         Top             =   660
         Width           =   3435
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3210
         Picture         =   "CommandForm.frx":20AE
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2205
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   480
      Top             =   4290
   End
   Begin VB.Timer TimerLocale 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   930
      Top             =   4290
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   400
      Left            =   9870
      TabIndex        =   21
      Top             =   8100
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   410
      Left            =   1020
      TabIndex        =   23
      Top             =   7710
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   420
      Left            =   3810
      TabIndex        =   24
      Top             =   5790
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   411
      Left            =   3810
      TabIndex        =   25
      Top             =   7140
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   425
      Left            =   13950
      TabIndex        =   84
      Top             =   4800
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   426
      Left            =   6780
      TabIndex        =   88
      Top             =   1290
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   2355
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Strorage 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   2
      Left            =   13050
      TabIndex        =   104
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Strorage 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   1
      Left            =   2310
      TabIndex        =   103
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label Cesta 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   100
      Top             =   9240
      Width           =   1275
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Scrapt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Index           =   0
      Left            =   3960
      TabIndex        =   99
      Top             =   9270
      Width           =   1245
   End
   Begin VB.Shape SeparTrasp 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   915
      Index           =   5
      Left            =   5700
      Top             =   2490
      Width           =   225
   End
   Begin VB.Shape SeparTrasp 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   915
      Index           =   4
      Left            =   10920
      Top             =   2490
      Width           =   225
   End
   Begin VB.Shape SeparTrasp 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   3
      Left            =   11730
      Top             =   4860
      Width           =   225
   End
   Begin VB.Shape SeparTrasp 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1005
      Index           =   2
      Left            =   9990
      Top             =   4860
      Width           =   225
   End
   Begin VB.Shape SeparTrasp 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   1
      Left            =   7800
      Top             =   4830
      Width           =   225
   End
   Begin VB.Shape SeparTrasp 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Index           =   0
      Left            =   5430
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "S 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   4230
      TabIndex        =   95
      Top             =   2130
      Width           =   405
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "S 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   9480
      TabIndex        =   94
      Top             =   2160
      Width           =   435
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "S 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   12300
      TabIndex        =   93
      Top             =   2160
      Width           =   405
   End
   Begin VB.Label Destination 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   2490
      TabIndex        =   92
      Top             =   2730
      Width           =   1605
   End
   Begin VB.Label Destination 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   12
      Left            =   7650
      TabIndex        =   91
      Top             =   2730
      Width           =   1605
   End
   Begin VB.Label Destination 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   12510
      TabIndex        =   90
      Top             =   2790
      Width           =   1605
   End
   Begin VB.Label Destination 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   12450
      TabIndex        =   89
      Top             =   3540
      Width           =   1605
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   11010
      TabIndex        =   83
      Top             =   5730
      Width           =   570
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   16
      Left            =   10980
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   645
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   15
      Left            =   11010
      TabIndex        =   82
      Top             =   5250
      Width           =   570
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   15
      Left            =   10980
      Shape           =   3  'Circle
      Top             =   5130
      Width           =   645
   End
   Begin VB.Shape Reggiat_2_1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   11070
      Top             =   5850
      Width           =   465
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   14
      Left            =   11010
      TabIndex        =   81
      Top             =   4710
      Width           =   570
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   14
      Left            =   10980
      Shape           =   3  'Circle
      Top             =   4620
      Width           =   645
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   13
      Left            =   10380
      TabIndex        =   80
      Top             =   5730
      Width           =   570
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   13
      Left            =   10350
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   645
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   6600
      TabIndex        =   52
      Top             =   8760
      Width           =   375
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   12
      Left            =   10380
      TabIndex        =   79
      Top             =   5220
      Width           =   570
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   12
      Left            =   10350
      Shape           =   3  'Circle
      Top             =   5130
      Width           =   645
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   10380
      TabIndex        =   78
      Top             =   4680
      Width           =   570
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   11
      Left            =   10350
      Shape           =   3  'Circle
      Top             =   4620
      Width           =   645
   End
   Begin VB.Shape Reggiatr_1_2 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   10590
      Top             =   4590
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1305
      Left            =   11220
      Top             =   4590
      Width           =   165
   End
   Begin VB.Shape Reggiatr_1_1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   10440
      Top             =   5850
      Width           =   465
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1185
      Index           =   411
      Left            =   2370
      Top             =   7140
      Width           =   2865
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   480
      Left            =   12390
      TabIndex        =   86
      Top             =   5040
      Width           =   345
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Side conveyors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   422
      Left            =   12720
      TabIndex        =   48
      Top             =   5100
      Width           =   1545
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   486
      Left            =   12390
      TabIndex        =   27
      Top             =   4200
      Width           =   345
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   4
      Left            =   12360
      Top             =   5010
      Width           =   2085
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Strorage 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Index           =   426
      Left            =   7680
      TabIndex        =   85
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   6
      Left            =   7500
      Top             =   1590
      Width           =   1905
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   885
      Index           =   426
      Left            =   1710
      Top             =   2490
      Width           =   13515
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   425
      Left            =   12840
      TabIndex        =   46
      Top             =   4350
      Width           =   1485
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   5
      Left            =   12360
      Top             =   4170
      Width           =   2205
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   8130
      TabIndex        =   77
      Top             =   8400
      Width           =   570
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   3510
      TabIndex        =   76
      Top             =   9960
      Width           =   390
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   4260
      TabIndex        =   75
      Top             =   1410
      Width           =   390
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   9510
      TabIndex        =   74
      Top             =   1410
      Width           =   390
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   12300
      TabIndex        =   73
      Top             =   1410
      Width           =   390
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   5610
      TabIndex        =   72
      Top             =   1560
      Width           =   435
   End
   Begin VB.Label EmNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   10860
      TabIndex        =   33
      Top             =   1500
      Width           =   390
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   270
      TabIndex        =   32
      Top             =   7710
      Width           =   435
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   11490
      TabIndex        =   31
      Top             =   6930
      Width           =   435
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   10
      Left            =   7050
      Shape           =   3  'Circle
      Top             =   8520
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   9
      Left            =   3390
      Shape           =   3  'Circle
      Top             =   9900
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   8
      Left            =   4140
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   7
      Left            =   9390
      Shape           =   3  'Circle
      Top             =   1320
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   6
      Left            =   12180
      Shape           =   3  'Circle
      Top             =   1350
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   5
      Left            =   5490
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   4
      Left            =   10740
      Shape           =   3  'Circle
      Top             =   1410
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   3
      Left            =   11370
      Shape           =   3  'Circle
      Top             =   6870
      Width           =   645
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   2
      Left            =   150
      Shape           =   3  'Circle
      Top             =   7590
      Width           =   645
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   32
      Left            =   6750
      Top             =   8430
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   31
      Left            =   7800
      Top             =   8700
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   30
      Left            =   7680
      Top             =   8460
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   29
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   28
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   27
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   26
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   25
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   24
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   32
      Left            =   8370
      TabIndex        =   71
      Top             =   8490
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   31
      Left            =   7230
      TabIndex        =   70
      Top             =   8550
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   30
      Left            =   6990
      TabIndex        =   69
      Top             =   8670
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   29
      Left            =   8070
      TabIndex        =   68
      Top             =   8670
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   28
      Left            =   7920
      TabIndex        =   67
      Top             =   8760
      Width           =   345
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   27
      Left            =   7770
      TabIndex        =   66
      Top             =   8430
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   26
      Left            =   7710
      TabIndex        =   65
      Top             =   8310
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   25
      Left            =   7710
      TabIndex        =   64
      Top             =   8310
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "2423"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   24
      Left            =   7710
      TabIndex        =   63
      Top             =   8310
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   23
      Left            =   6960
      TabIndex        =   62
      Top             =   8400
      Width           =   375
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   23
      Left            =   7710
      Top             =   8310
      Width           =   165
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   22
      Left            =   6570
      TabIndex        =   61
      Top             =   8310
      Width           =   375
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   195
      Index           =   22
      Left            =   6120
      Top             =   8610
      Width           =   855
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   21
      Left            =   7410
      TabIndex        =   60
      Top             =   8430
      Width           =   375
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   21
      Left            =   7350
      Top             =   8400
      Width           =   165
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   20
      Left            =   7560
      TabIndex        =   59
      Top             =   8610
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   19
      Left            =   8370
      TabIndex        =   58
      Top             =   8550
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   7830
      TabIndex        =   57
      Top             =   8430
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   7860
      TabIndex        =   56
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   7200
      TabIndex        =   55
      Top             =   8310
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   8310
      TabIndex        =   54
      Top             =   8490
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   7770
      TabIndex        =   53
      Top             =   8550
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   6690
      TabIndex        =   51
      Top             =   8430
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   1020
      TabIndex        =   50
      Top             =   2340
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   8040
      TabIndex        =   49
      Top             =   8730
      Width           =   375
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   20
      Left            =   8640
      Top             =   8280
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   19
      Left            =   8280
      Top             =   8310
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   195
      Index           =   18
      Left            =   7830
      Top             =   8730
      Width           =   555
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   17
      Left            =   8430
      Top             =   8280
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   16
      Left            =   8430
      Top             =   8340
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   15
      Left            =   8130
      Top             =   8430
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   14
      Left            =   7890
      Top             =   8460
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   13
      Left            =   7590
      Top             =   8340
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   12
      Left            =   6570
      Top             =   8490
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   11
      Left            =   7950
      Top             =   8670
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   10
      Left            =   7680
      Top             =   8640
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   9
      Left            =   7350
      Top             =   8670
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   8
      Left            =   1200
      Top             =   2220
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   165
      Index           =   7
      Left            =   5520
      Top             =   2280
      Width           =   585
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   165
      Index           =   6
      Left            =   10770
      Top             =   2280
      Width           =   585
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   4
      Left            =   5340
      Top             =   3660
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   3
      Left            =   5550
      Top             =   7200
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   2
      Left            =   1860
      Top             =   6510
      Width           =   165
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   39
      Top             =   10080
      Width           =   255
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   165
      Index           =   1
      Left            =   4080
      Top             =   9930
      Width           =   585
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   450
      Left            =   3180
      TabIndex        =   19
      Top             =   8490
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   14730
      Picture         =   "CommandForm.frx":413C
      Stretch         =   -1  'True
      Top             =   9420
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   14730
      Picture         =   "CommandForm.frx":457E
      Stretch         =   -1  'True
      Top             =   8850
      Width           =   540
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1575
      Index           =   425
      Left            =   12270
      Top             =   3360
      Width           =   2955
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "MPS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   420
      Left            =   2910
      TabIndex        =   47
      Top             =   6300
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Legend"
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
      Height          =   315
      Left            =   12090
      TabIndex        =   45
      Top             =   8430
      Width           =   3105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   0
      Left            =   12030
      Top             =   8340
      Width           =   3285
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   165
      Index           =   5
      Left            =   10680
      Top             =   6660
      Width           =   585
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   5730
      TabIndex        =   44
      Top             =   7290
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   10860
      TabIndex        =   43
      Top             =   6810
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "SB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   12330
      TabIndex        =   41
      Top             =   8940
      Width           =   405
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   5490
      TabIndex        =   40
      Top             =   3750
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   38
      Top             =   6630
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   5730
      TabIndex        =   37
      Top             =   2010
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   7680
      TabIndex        =   36
      Top             =   8610
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   8400
      TabIndex        =   35
      Top             =   8430
      Width           =   375
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   6120
      TabIndex        =   34
      Top             =   6780
      Width           =   435
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "PE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   12150
      TabIndex        =   30
      Top             =   9540
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   10980
      TabIndex        =   28
      Top             =   1980
      Width           =   255
   End
   Begin VB.Image Armadio 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   9450
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   470
      Left            =   2760
      TabIndex        =   26
      Top             =   6210
      Width           =   345
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Rollway"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Index           =   410
      Left            =   3420
      TabIndex        =   20
      Top             =   8460
      Width           =   1545
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   1
      Left            =   6000
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   645
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   0
      Left            =   12090
      Top             =   8850
      Width           =   225
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   0
      Left            =   12030
      Shape           =   3  'Circle
      Top             =   9450
      Width           =   645
   End
   Begin VB.Label Commento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Safety door number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   12810
      TabIndex        =   42
      Top             =   9000
      Width           =   1890
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pulsanti emergenza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   12840
      TabIndex        =   29
      Top             =   9540
      Width           =   1905
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1185
      Index           =   10
      Left            =   12030
      Top             =   8790
      Width           =   3285
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   555
      Index           =   3
      Left            =   2730
      Top             =   6180
      Width           =   2205
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1365
      Index           =   420
      Left            =   2370
      Top             =   5790
      Width           =   2865
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   525
      Index           =   1
      Left            =   2910
      Shape           =   2  'Oval
      Top             =   8340
      Width           =   2205
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   585
      Index           =   410
      Left            =   120
      Top             =   8310
      Width           =   5115
   End
   Begin VB.Shape Stoccaggio 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1095
      Index           =   2
      Left            =   6990
      Top             =   1410
      Width           =   2955
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   885
      Index           =   422
      Left            =   2370
      Top             =   4920
      Width           =   12855
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   675
      Index           =   0
      Left            =   3990
      Top             =   9090
      Width           =   1125
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1065
      Index           =   0
      Left            =   2370
      Top             =   8910
      Width           =   2865
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   2
      Left            =   2130
      Top             =   1590
      Width           =   1905
   End
   Begin VB.Shape Stoccaggio 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1095
      Index           =   3
      Left            =   1710
      Top             =   1410
      Width           =   2955
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   7
      Left            =   12870
      Top             =   1590
      Width           =   1905
   End
   Begin VB.Shape Stoccaggio 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1095
      Index           =   1
      Left            =   12270
      Top             =   1410
      Width           =   2955
   End
End
Attribute VB_Name = "CommandForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Storage As Integer
Public PresentazioneOff As Boolean
Private Barriera(32) As Boolean
Private Emergenza(16) As Boolean

Private Sub AdjRolls_Click(Index As Integer)
   Select Case Index
    Case 0
       If AdjRolls(0).value Then
            DB422.Bit(26, 5) = 1
            DB422.Bit(26, 6) = 0
            DB422.Bit(26, 7) = 0
            AdjRolls(1).value = 0
            AdjRolls(2).value = 0
       Else
            DB422.Bit(26, 5) = False
       End If
    Case 1
        If AdjRolls(1).value Then
            DB422.Bit(26, 5) = 0
            DB422.Bit(26, 6) = 1
            DB422.Bit(26, 7) = 0
            AdjRolls(0).value = 0
            AdjRolls(2).value = 0
        Else
            DB422.Bit(26, 6) = 0
        End If
    Case 2
        If AdjRolls(2).value Then
            DB422.Bit(26, 5) = 0
            DB422.Bit(26, 6) = 0
            DB422.Bit(26, 7) = 1
            AdjRolls(0).value = 0
            AdjRolls(1).value = 0
        Else
            DB422.Bit(26, 7) = 0
        End If
    End Select
End Sub

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub LblRifVel_Click(Index As Integer)
            TOUCHNumericPad.ValoreMin = 20
            TOUCHNumericPad.ValoreMax = 100
            TOUCHNumericPad.Dati = DB450.Word(24)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                DB450.Word(24) = TOUCHNumericPad.Dati
            End If
End Sub

Private Sub Cesta_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 0
    TOUCHNumericPad.Dati = DB410.Word(38)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        DB410.Word(38) = TOUCHNumericPad.Dati
    End If
End Sub


'Private Sub ComboDest_Click(Index As Integer)
'  Select Case Index
'  Case 0
'        DB426.Word(68) = ComboDest(Index).ListIndex + 1
'  Case 1
'        DB426.Word(66) = ComboDest(Index).ListIndex + 1
'  Case 2
'        DB426.Word(64) = ComboDest(Index).ListIndex + 1
'  Case 3
'        DB426.Word(62) = ComboDest(Index).ListIndex + 1
'  End Select
'End Sub

Private Sub ComboStorage_Click()
   Select Case ComboStorage.ListIndex
    Case 0
        DB426.Word(70) = 0
    Case 1
       DB426.Word(70) = 1
   Case 2
       DB426.Word(70) = 2
   Case 3
        DB426.Word(70) = 3
   Case 4
       DB426.Word(70) = 4
    Case Else
       DB426.Word(70) = 0
   End Select
End Sub

Private Sub Destination_Click(Index As Integer)
   TOUCHNumericPad.Decimali = 0
    Select Case Index
    Case 10
         '   TOUCHNumericPad.ValoreMin = 1
         '   TOUCHNumericPad.ValoreMax = 3
         '   TOUCHNumericPad.Dati = DB426.Word(62)
         '   TOUCHNumericPad.Show vbModal
         '   If TOUCHNumericPad.DatiConfermati Then
         '       DB426.Word(62) = TOUCHNumericPad.Dati
         '   End If
    Case 11
            TOUCHNumericPad.ValoreMin = 1
            TOUCHNumericPad.ValoreMax = 3
            TOUCHNumericPad.Dati = DB426.Word(64)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                DB426.Word(64) = TOUCHNumericPad.Dati
            End If
    Case 12
            TOUCHNumericPad.ValoreMin = 1
            TOUCHNumericPad.ValoreMax = 3
            TOUCHNumericPad.Dati = DB426.Word(66)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                DB426.Word(66) = TOUCHNumericPad.Dati
            End If
    Case 13
            TOUCHNumericPad.ValoreMin = 1
            TOUCHNumericPad.ValoreMax = 3
            TOUCHNumericPad.Dati = DB426.Word(68)
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                DB426.Word(68) = TOUCHNumericPad.Dati
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
   Dim i As Integer
    ' LblRifVel(1) = DB450.Word(24)
   DB400.WORDSReadAsync 18, 3
   TestBarriere
End Sub

'Private Sub Timer1_Timer()

'End Sub

'Private Sub Timer1_Timer()
'Dim i
'   List1.Clear
'   DB470.BlockReadAsync
'   For i = 1 To DB470.NumItems
'     List1.AddItem DB470.MultiReadValore(i)
'  Next
'   DB480.BlockReadAsync
'   For i = 1 To DB480.NumItems
'     List1.AddItem DB480.MultiReadValore(i)
'  Next
'  DB486.BlockReadAsync
'   For i = 1 To DB486.NumItems
'     List1.AddItem DB486.MultiReadValore(i)
'  Next
'End Sub

'=======================================================================
' CAMBIA IL VALORE DELLA VARIABILE CHE INDICA AL KENERL LA PRESENZA DI ALLARMI
'=======================================================================

Private Sub TimerLocale_Timer()
    Static oneShot As Integer
    Dim AllOn As Boolean
    Dim i As Integer
    Dim p As Boolean
    Dim ctl As Control
    Dim a As Integer
    
    On Error Resume Next
   
    For i = Allarme1.LBound To Allarme1.UBound
       If Allarme1(i).AllarmeTipo <> Nessuno Then Allarme1(i).RefreshTimer
    Next
    Me.Update
    On Error GoTo 0
    
   ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
  
End Sub
   
'==================================================================
' FUNZIONE DI AGGIORNAMENTO DELLO STATO DELLA PAGINA
'==================================================================

Public Sub Update()
   Dim i As Integer
   
'   Static dest0_old As Integer
'   Static dest1_old As Integer
'   Static dest2_old As Integer
'   Static dest3_old As Integer
'
   Cesta.caption = DB410.Word(38)  ' aggiorna numero tubi su cesta scarti

'==================================================================
' VISUALIZZA LE DESTINAZIONI PACCHI
'==================================================================

   Destination(10).Visible = DB426.Word(62) <> 0
'   ComboDest(3).Visible = DB426.Word(62) <> 0
'
'   If dest3_old <> Abs(DB426.Word(62) - 1) Then
'      ComboDest(3).ListIndex = Abs(DB426.Word(62) - 1)
'      dest3_old = Abs(DB426.Word(62) - 1)
'   End If
'   If dest2_old <> Abs(DB426.Word(64) - 1) Then
'      ComboDest(2).ListIndex = Abs(DB426.Word(64) - 1)
'      dest2_old = Abs(DB426.Word(64) - 1)
'   End If
'   If dest1_old <> Abs(DB426.Word(66) - 1) Then
'      ComboDest(1).ListIndex = Abs(DB426.Word(66) - 1)
'      dest1_old = Abs(DB426.Word(66) - 1)
'   End If
'   If dest0_old <> Abs(DB426.Word(68) - 1) Then
'      ComboDest(0).ListIndex = Abs(DB426.Word(68) - 1)
'      dest0_old = Abs(DB426.Word(68) - 1)
'   End If

   Destination(10) = "S" & Str(DB426.Word(62))

   If DB426.Word(64) Then
   Destination(11) = "S" & Str(DB426.Word(64))
   Else
   Destination(11) = ""
   End If

   If DB426.Word(66) Then
   Destination(12) = "S" & Str(DB426.Word(66))
   Else
   Destination(12) = ""
   End If

   If DB426.Word(68) Then
   Destination(13) = "S" & Str(DB426.Word(68))
   Else
   Destination(13) = ""
   End If
  
'==================================================================
' VISUALIZZA GLI ALLARMI
'==================================================================
    
    'led di allarme generale

     AllarmiLayout DB410, 410
     AllarmiLayout DB411, 411
     AllarmiLayout DB400, 400
  '   AllarmiLayout DB412, 412
 '    AllarmiLayout DB413, 413
  '   AllarmiLayout DB414, 414
   '  AllarmiLayout DB415, 415
    ' AllarmiLayout DB416, 416
     'AllarmiLayout DB417, 417
     AllarmiLayout DB420, 420
    ' AllarmiLayout DB424, 424
     AllarmiLayout DB422, 422
     AllarmiLayout DB425, 425
     AllarmiLayout DB426, 426

'     AllarmiLayout DB419, manicottatrice
'     AllarmiLayout DB418, tappatrice
''
'==================================================================
' VISUALIZZA LO STATO DELLE ZONE
'==================================================================

     ColoreLayout DB410, 410
     ColoreLayout DB411, 411
    ' ColoreLayout DB415, 415
     ColoreLayout DB420, 420
     ColoreLayout DB422, 422
     ColoreLayout DB425, 425
     ColoreLayout DB426, 426
   '  ColoreLayout DB424, 424
    ' ColoreLayout DB416, 416
     'ColoreLayout DB417, 417
     
'==================================================================
' VISUALIZZA STATO BYPASS e filettatrici
'==================================================================
    
    
    'filettatrici

  '  If DB460.Bit(78, 0) = True Then
  '       Zona(7).Visible = True
  '       Zona(3).Visible = False
  '  Else
  '       Zona(7).Visible = False
  '       Zona(3).Visible = True
  '  End If
'
'    If DB460.Bit(118, 0) = True Then
'         Zona(4).Visible = False
'         Zona(8).Visible = True
'    Else
'         Zona(4).Visible = True
'         Zona(8).Visible = False
'    End If

'==================================================================
' VISUALIZZA L'ORDINE
'==================================================================
   
    LblCodOrd(470).caption = frmKernel.CodOrdineCorrente.CodPacco
    LblCodOrd(480).caption = frmKernel.CodOrdineCorrente.CodRegge
    LblCodOrd(486).caption = frmKernel.CodOrdineCorrente.CodStoccaggio
  '  LblCodOrd(460).Caption = frmKernel.CodOrdineCorrente.CodWB
    LblCodOrd(450).caption = frmKernel.CodOrdineCorrente.CodEntrata
    
End Sub


''==================================================================
'' FUNZIONE ABILITAZIONE ALLARMI
''==================================================================

Private Sub AllarmiLayout(DBSource As DBClass, Index As Integer)
    If DBSource.MaskBit(0, 8) = True And DBSource.MaskBit(0, 9) = True Then
       Allarme1(Index).AllarmeTipo = Entrambi
    Else
        If DBSource.MaskBit(0, 8) = True And DBSource.MaskBit(0, 9) = False Then
            Allarme1(Index).AllarmeTipo = Allarme
        Else
            If DBSource.MaskBit(0, 8) = False And DBSource.MaskBit(0, 9) = True Then
                 Allarme1(Index).AllarmeTipo = Messaggio
            Else
                If DBSource.MaskBit(0, 8) = False And DBSource.MaskBit(0, 9) = False Then
                    Allarme1(Index).AllarmeTipo = Nessuno
                End If
            End If
        End If
    End If
End Sub

'==================================================================
' FUNZIONE COLORAZIONE LAYOUT
'==================================================================

Private Sub ColoreLayout(DBSource As DBClass, Index As Integer)
    If DBSource.MaskBit(0, 2) Then
        LabelStato(Index).BackColor = ManualColor
    Else
        If DBSource.MaskBit(0, 3) Then
            LabelStato(Index).BackColor = SemiautoColor
        Else
            If DBSource.MaskBit(0, 4) Then
                LabelStato(Index).BackColor = AutoColor
            Else
                LabelStato(Index).BackColor = E_StopColor
            End If
        End If
    End If
    
    Stoccaggio(1).BackColor = LabelStato(426).BackColor
    Stoccaggio(2).BackColor = LabelStato(426).BackColor
    Stoccaggio(3).BackColor = LabelStato(426).BackColor
End Sub

'==================================================================
' EVENTO FOCUS AL FORM
'==================================================================

Public Sub Form_Activate()
 ' Barra1.Pulsante_Click 3
  Barra21.Selezionato = 3
  
  ComboStorage.Clear
  ComboStorage.AddItem "Auto"
  ComboStorage.AddItem "Storage 1"
  ComboStorage.AddItem "Storage 2"
  ComboStorage.AddItem "Storage 3"
  ComboStorage.AddItem "From recipe"
  
  Storage = Abs(DB426.Word(70))
  ComboStorage.ListIndex = Storage
   
  If PresentazioneOff = False Then
    TimerLocale.Enabled = False
    Timer1.Enabled = False
  '  Barra1.Bloccata = True
  Else
    Call Aggiornamento
  End If
 If DB422.Bit(26, 5) = True Then
    AdjRolls(0).value = 1
 End If
 If DB422.Bit(26, 6) = True Then
     AdjRolls(1).value = 1
 End If
 If DB422.Bit(26, 7) = True Then
    AdjRolls(2).value = 1
 End If
  lblbar(2) = Trim(Param.GetNumber("Par111_Password utente"))
End Sub
Sub Aggiornamento()
    frmKernel.PulAllarmiPremuto = False
    TimerLocale.Enabled = True
    Timer1.Enabled = True
    Timer1.Interval = 500
    TimerLocale.Interval = 500
    Me.Update
 '   Barra1.Bloccata = False
    TestBarriere
End Sub
Private Sub Form_Deactivate()
    TimerLocale.Enabled = False
    Timer1.Enabled = False
End Sub

'==================================================================
' EVENTO CARICA RISORSE NEL FORM
'==================================================================

Private Sub Form_Load()
Dim i As Integer

    WindowState = 2
   
    On Error Resume Next
    For i = Allarme1.LBound To Allarme1.UBound
       Allarme1(i).AllarmeTipo = Nessuno
       Allarme1(i).Intervallo = 500
    Next
    On Error GoTo 0
   Armadio = LoadPicture("..\bitmap\ArmadioElettrico.gif")

   ScritteMultilingua
End Sub

'==================================================================
' VISUALIZZA LA PAGINA DEGLI ALLARMI IN BASE ALL'ALLARME CLICCATO
'==================================================================

Private Sub Allarme1_Cliccato(Index As Integer)
     
    frmKernel.PulAllarmiPremuto = True
    'inizializzazione allarmi

    AlarmForm.CheckDB400.value = 0
    AlarmForm.CheckDB410.value = 0
    AlarmForm.CheckDB411.value = 0
'    AlarmForm.CheckDB412.value = 0
'    AlarmForm.CheckDB413.Value = 0
'    AlarmForm.CheckDB414.value = 0
'    AlarmForm.CheckDB415.value = 0
'    AlarmForm.CheckDB416.value = 0
'    AlarmForm.CheckDB417.value = 0
'    AlarmForm.CheckDB418.value = 0
'    AlarmForm.CheckDB419.value = 0
    AlarmForm.CheckDB420.value = 0
    AlarmForm.CheckDB422.value = 0
 '   AlarmForm.CheckDB424.value = 0
    AlarmForm.CheckDB425.value = 0
    AlarmForm.CheckDB426.value = 0
    
   'Setta l'allarme selezionato
   
    Select Case Index
        Case 400
            AlarmForm.CheckDB400.value = 1
        Case 410
            AlarmForm.CheckDB410.value = 1
        Case 411
            AlarmForm.CheckDB411.value = 1
        Case 412
            AlarmForm.CheckDB412.value = 1
        Case 413
            AlarmForm.CheckDB413.value = 1
        Case 414
            AlarmForm.CheckDB414.value = 1
        Case 415
            AlarmForm.CheckDB415.value = 1
        Case 416
            AlarmForm.CheckDB416.value = 1
        Case 417
            AlarmForm.CheckDB417.value = 1
        Case 418
            AlarmForm.CheckDB418.value = 1
        Case 419
            AlarmForm.CheckDB419.value = 1
        Case 420
            AlarmForm.CheckDB420.value = 1
        Case 422
            AlarmForm.CheckDB422.value = 1
        Case 424
            AlarmForm.CheckDB424.value = 1
        Case 425
            AlarmForm.CheckDB425.value = 1
        Case 426
            AlarmForm.CheckDB426.value = 1
    End Select
    
    'visualizza il form allarmi
    
    AlarmForm.ZOrder 0
    AlarmForm.Visible = True
    AlarmForm.WindowState = vbMaximized
    
End Sub

Private Sub Label5_Click()
   ErrDBPiccolo = False
End Sub

Sub ScritteMultilingua()
   lblbar(0) = Param.Text("Impianto")
   lblbar(2) = Trim(Param.GetNumber("Par111_Password utente")) ' Param.Text("Comm_Num")
   lblbar(4) = Param.Text("Anno")
   lblbar(5) = Param.Text("Comm_Anno")
   lblbar(1) = Param.Text("Pagina")
   Commento.caption = Param.Text("LegBarr")
   Label1 = Param.Text("PulEm")
   lblZoneName(410) = Param.Text("ViaRulli")
   lblZoneName(420) = Param.Text("PP")
'   lblZoneName(422) = Param.Text("Trlat")
   lblZoneName(425) = Param.Text("Trasferimento")
'   lblZoneName(426) = Param.Text("Reggiatura")
'   lblZoneName(426) = Param.Text("Stoccaggio")
   lblZoneName(0) = Param.Text("Scarti")
   
   ComboDest(0).Clear
   ComboDest(0).AddItem "St. 1"
   ComboDest(0).AddItem "St. 2"
   ComboDest(0).AddItem "St. 3"
   ComboDest(1).Clear
   ComboDest(1).AddItem "St. 1"
   ComboDest(1).AddItem "St. 2"
   ComboDest(1).AddItem "St. 3"
   ComboDest(2).Clear
   ComboDest(2).AddItem "St. 1"
   ComboDest(2).AddItem "St. 2"
   ComboDest(2).AddItem "St. 3"
   ComboDest(3).Clear
   ComboDest(3).AddItem "St. 1"
   ComboDest(3).AddItem "St. 2"
   ComboDest(3).AddItem "St. 3"
End Sub

Sub TestBarriere()
    Dim i As Integer
    Dim Br As Boolean
    Dim Em As Boolean
    Static Lamp As Boolean
    
    Lamp = Not Lamp
    On Error Resume Next
    DoEvents
    
    Br = False: Em = False
    For i = 0 To 31
        BarrPallino(i + 1).Visible = (DB400.MaskBit(20 + 2 * Abs(i > 15), i - 16 * Abs(i > 15))) And Lamp
        BarrNumero(i + 1).Visible = BarrPallino(i + 1).Visible
        Br = Br Or BarrPallino(i + 1).Visible
    Next

    For i = 0 To 15
       Emergenza(i + 1) = DB400.MaskBit(18, i)
       EmPallino(i + 1).Visible = Emergenza(i + 1) And Lamp
       EmNumero(i + 1).Visible = EmPallino(i + 1).Visible
       Em = Em Or EmPallino(i + 1).Visible
    Next
    Image1.Visible = Br
    Image4.Visible = Em

End Sub

Private Sub XPButton1_Click(Index As Integer)
  Select Case Index
     Case 0
            frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
            frmStatistica.Show
            frmConnecting.Hide
     Case 1
          ' frmConnecting.ShowConnecting "Refreshing param grid. Please wait...", Me
           Param.ChiamataService = True
           frmKernel.PaginaCorrente = PagService
          ' frmConnecting.Hide
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
                .NomeFile = "mappa_pagina.htm"
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

