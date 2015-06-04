VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form BundleForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   4785
      Left            =   1950
      TabIndex        =   59
      Top             =   10140
      Visible         =   0   'False
      Width           =   5355
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   2820
         Top             =   2910
         Width           =   2115
      End
      Begin VB.Shape ShapeRif 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   660
         Top             =   2910
         Width           =   2115
      End
      Begin VB.Image SelettoreSagome 
         Height          =   1275
         Left            =   3210
         Top             =   3180
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Image SelettorePale 
         Height          =   1155
         Left            =   1140
         Top             =   3210
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sagome hex"
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
         Height          =   510
         Left            =   2820
         TabIndex        =   65
         Top             =   2400
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PALE VELOCI"
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
         Height          =   510
         Left            =   660
         TabIndex        =   64
         Top             =   2400
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0          1"
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
         Left            =   3030
         TabIndex        =   63
         Top             =   2850
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0          1"
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
         Left            =   930
         TabIndex        =   62
         Top             =   2880
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alternato     Continuo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   210
         TabIndex        =   61
         Top             =   330
         Width           =   3105
      End
      Begin VB.Label LblScivolo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scivolo"
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
         Height          =   390
         Left            =   1200
         TabIndex        =   60
         Top             =   1800
         Width           =   3510
      End
      Begin VB.Image SelScivolo 
         Height          =   1155
         Left            =   1470
         Top             =   600
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2745
      Left            =   30
      TabIndex        =   28
      Top             =   7560
      Width           =   4455
      Begin dp6.ControlloPacco DisegnoTubo 
         Height          =   2865
         Left            =   120
         TabIndex        =   29
         Top             =   -240
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   5054
      End
      Begin VB.Image ImgLung 
         Height          =   795
         Left            =   2370
         Picture         =   "BundleForm.frx":0000
         Top             =   1500
         Width           =   2040
      End
      Begin VB.Label LblLung 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2760
         TabIndex        =   46
         Top             =   1140
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   6405
      Left            =   60
      TabIndex        =   23
      Top             =   1140
      Width           =   4425
      Begin VB.CommandButton cmdTubesChains 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   2
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1410
         Width           =   945
      End
      Begin VB.CommandButton cmdTubesChains 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   3
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1410
         Width           =   945
      End
      Begin VB.CommandButton cmdTubesChains 
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   600
         Width           =   945
      End
      Begin VB.CommandButton cmdTubesChains 
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   1
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   600
         Width           =   945
      End
      Begin dp6.ControlloUpDown ControlloTempoAllineamentoTubo 
         Height          =   990
         Left            =   630
         TabIndex        =   25
         Top             =   3840
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1746
      End
      Begin dp6.ControlloUpDown ControlloMagneti 
         Height          =   945
         Left            =   600
         TabIndex        =   24
         Top             =   5370
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1667
      End
      Begin VB.Label TubesLastRow 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1800
         TabIndex        =   78
         Top             =   2730
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblLastRow 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tubes for last row bundle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   90
         TabIndex        =   77
         Top             =   2310
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Side    exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   1140
         TabIndex        =   76
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Side    entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1140
         TabIndex        =   75
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3090
         TabIndex        =   74
         Top             =   1410
         Width           =   1275
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total tubes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3060
         TabIndex        =   73
         Top             =   600
         Width           =   1305
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tubes in the entry stack"
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
         Height          =   345
         Left            =   90
         TabIndex        =   72
         Top             =   180
         Width           =   4245
      End
      Begin VB.Label DisplayTubiAccum 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3300
         TabIndex        =   67
         Top             =   930
         Width           =   885
      End
      Begin VB.Label DisplayTubiAccum 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   3330
         TabIndex        =   66
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alignment time (s)"
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
         Height          =   345
         Left            =   90
         TabIndex        =   26
         Top             =   3450
         Width           =   4245
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vel. magneti (%)"
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
         Height          =   345
         Left            =   90
         TabIndex        =   27
         Top             =   4980
         Width           =   4260
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   9165
      Left            =   4500
      TabIndex        =   12
      Top             =   1140
      Width           =   10755
      Begin VB.CommandButton Magnetposition 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Magnet position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   8220
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6990
         Width           =   1875
      End
      Begin VB.CommandButton CommandModifica 
         BackColor       =   &H0000FFFF&
         Caption         =   "MODIFICA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8430
         Width           =   3045
      End
      Begin VB.Frame FrameMagnet 
         BackColor       =   &H00C0C0C0&
         Height          =   9165
         Left            =   210
         TabIndex        =   33
         Top             =   150
         Width           =   7515
         Begin VB.Label Magnetizzato 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   5940
            TabIndex        =   52
            Top             =   3120
            Width           =   945
         End
         Begin VB.Label Magnetizzato 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   4890
            TabIndex        =   51
            Top             =   3120
            Width           =   945
         End
         Begin VB.Label Magnetizzato 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   3840
            TabIndex        =   50
            Top             =   3120
            Width           =   945
         End
         Begin VB.Label Magnetizzato 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2790
            TabIndex        =   49
            Top             =   3090
            Width           =   945
         End
         Begin VB.Label Magnetizzato 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   48
            Top             =   3090
            Width           =   945
         End
         Begin VB.Label Magnetizzato 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   690
            TabIndex        =   47
            Top             =   3090
            Width           =   945
         End
         Begin VB.Image ImgAsseMag 
            Height          =   2085
            Index           =   1
            Left            =   2040
            Picture         =   "BundleForm.frx":01B3
            Stretch         =   -1  'True
            Top             =   1470
            Width           =   315
         End
         Begin VB.Image ImgAsseMag 
            Height          =   2115
            Index           =   0
            Left            =   990
            Picture         =   "BundleForm.frx":035E
            Stretch         =   -1  'True
            Top             =   1470
            Width           =   315
         End
         Begin VB.Image ImgAsseMag 
            Height          =   2085
            Index           =   5
            Left            =   6240
            Picture         =   "BundleForm.frx":0509
            Stretch         =   -1  'True
            Top             =   1500
            Width           =   315
         End
         Begin VB.Image ImgAsseMag 
            Height          =   2085
            Index           =   4
            Left            =   5190
            Picture         =   "BundleForm.frx":06B4
            Stretch         =   -1  'True
            Top             =   1500
            Width           =   315
         End
         Begin VB.Image ImgAsseMag 
            Height          =   2085
            Index           =   3
            Left            =   4140
            Picture         =   "BundleForm.frx":085F
            Stretch         =   -1  'True
            Top             =   1500
            Width           =   315
         End
         Begin VB.Image ImgAsseMag 
            Height          =   2085
            Index           =   2
            Left            =   3090
            Picture         =   "BundleForm.frx":0A0A
            Stretch         =   -1  'True
            Top             =   1470
            Width           =   315
         End
         Begin VB.Image ImageMagnete 
            Height          =   1935
            Index           =   3
            Left            =   3810
            Picture         =   "BundleForm.frx":0BB5
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Image ImageMagnete 
            Height          =   1935
            Index           =   1
            Left            =   1710
            Picture         =   "BundleForm.frx":0F11
            Stretch         =   -1  'True
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Enable"
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
            Height          =   510
            Index           =   3
            Left            =   690
            TabIndex        =   39
            Top             =   210
            Width           =   6210
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Z position [inch]"
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
            Height          =   510
            Index           =   1
            Left            =   660
            TabIndex        =   38
            Top             =   6510
            Width           =   6210
         End
         Begin VB.Label PosX 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   555
            Index           =   6
            Left            =   2760
            TabIndex        =   37
            Top             =   7110
            Width           =   1890
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "        1        2        3         4        5        6"
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
            Height          =   495
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   7185
         End
         Begin VB.Image ImageMagnete 
            Height          =   1935
            Index           =   5
            Left            =   5910
            Picture         =   "BundleForm.frx":126D
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Image ImageMagnete 
            Height          =   1935
            Index           =   4
            Left            =   4860
            Picture         =   "BundleForm.frx":15C9
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Image ImageMagnete 
            Height          =   1935
            Index           =   2
            Left            =   2760
            Picture         =   "BundleForm.frx":1925
            Stretch         =   -1  'True
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Image ImageMagnete 
            Height          =   1935
            Index           =   0
            Left            =   660
            Picture         =   "BundleForm.frx":1C81
            Stretch         =   -1  'True
            Top             =   1410
            Width           =   1005
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "X position [inch]"
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
            Height          =   510
            Index           =   0
            Left            =   660
            TabIndex        =   36
            Top             =   4860
            Width           =   6210
         End
         Begin VB.Label PosX 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   0
            Left            =   2760
            TabIndex        =   35
            Top             =   5460
            Width           =   1890
         End
         Begin VB.Label Enab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   3405
            Index           =   1
            Left            =   1710
            TabIndex        =   41
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Enab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   3405
            Index           =   0
            Left            =   660
            TabIndex        =   40
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Enab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   3405
            Index           =   2
            Left            =   2760
            TabIndex        =   42
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Enab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   3405
            Index           =   3
            Left            =   3810
            TabIndex        =   43
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Enab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   3405
            Index           =   4
            Left            =   4860
            TabIndex        =   44
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Enab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   3405
            Index           =   5
            Left            =   5910
            TabIndex        =   45
            Top             =   960
            Width           =   990
         End
      End
      Begin dp6.ControlloPacco DisegnoPacco 
         Height          =   8850
         Left            =   330
         TabIndex        =   14
         Top             =   210
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   15610
      End
      Begin VB.Label lblpeso 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32000000"
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
         Height          =   465
         Left            =   5100
         TabIndex        =   58
         Top             =   8460
         Width           =   1665
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight [Kg]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5340
         TabIndex        =   57
         Top             =   8580
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order tubes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   56
         Top             =   8310
         Width           =   1605
      End
      Begin VB.Label lblTubesTot 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32000000"
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
         Height          =   495
         Left            =   5160
         TabIndex        =   55
         Top             =   8400
         Width           =   1665
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Totals MT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5160
         TabIndex        =   54
         Top             =   8610
         Width           =   1065
      End
      Begin VB.Label lblMTtot 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32000000"
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
         Height          =   465
         Left            =   5130
         TabIndex        =   53
         Top             =   9180
         Width           =   1665
      End
      Begin VB.Label BundleTubesCounterDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1500
         Left            =   7680
         TabIndex        =   22
         Top             =   870
         Width           =   2700
      End
      Begin VB.Label BundlesCounterDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "999"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1065
         Left            =   7710
         TabIndex        =   21
         Top             =   4440
         Width           =   2700
      End
      Begin VB.Label BundleTubesPresetDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   8490
         TabIndex        =   20
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label OfLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "di"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   8820
         TabIndex        =   19
         Top             =   2430
         Width           =   540
      End
      Begin VB.Label CurrentTubesLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tubi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   660
         Left            =   7680
         TabIndex        =   17
         Top             =   210
         Width           =   2715
      End
      Begin VB.Label BundlesPresetDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "26"
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
         Height          =   705
         Left            =   8490
         TabIndex        =   16
         Top             =   5910
         Width           =   1215
      End
      Begin VB.Label OfLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "di"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8790
         TabIndex        =   15
         Top             =   5520
         Width           =   645
      End
      Begin VB.Label CurrentBundleLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pacco n."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   810
         Left            =   7710
         TabIndex        =   18
         Top             =   3720
         Width           =   2685
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   7590
         TabIndex        =   30
         Top             =   3630
         Width           =   2925
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3390
         Left            =   7590
         TabIndex        =   31
         Top             =   180
         Width           =   2925
      End
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
         Index           =   2
         Left            =   12270
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   32
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
         Picture         =   "BundleForm.frx":1FDD
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
         TabIndex        =   10
         Top             =   270
         Width           =   1515
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
         TabIndex        =   9
         Top             =   690
         Width           =   2415
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
         TabIndex        =   8
         Top             =   240
         Width           =   2175
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
         TabIndex        =   7
         Top             =   270
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
         TabIndex        =   6
         Top             =   240
         Width           =   2415
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
         TabIndex        =   5
         Top             =   630
         Width           =   3495
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   2
         Left            =   3180
         TabIndex        =   4
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   4110
      Top             =   1020
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4530
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
            Picture         =   "BundleForm.frx":406B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BundleForm.frx":466E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
End
Attribute VB_Name = "BundleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tchange
        OldLu As Variant
        OldSp As Variant
        OldTi As Variant
        OldLa As Variant
        OldAl As Variant
End Type

Private PrimaScansione As Boolean
Private nTubi As Integer
Private PaginaAttiva As Boolean
Private NumProfilo As Integer
Private CodiceVecchio As Integer
Private MemOld(50) As Long
Private ChangeON As Boolean
Private TubeChange As Boolean
Private OldTube As tchange

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub cmdTubesChains_Click(Index As Integer)
 Select Case Index
    Case 0
          DB420.Bit(27, 1) = 1
    Case 1
          DB420.Bit(27, 0) = 1
    Case 2
          DB420.Bit(27, 3) = 1
    Case 3
          DB420.Bit(27, 2) = 1
    End Select
End Sub


Private Sub Magnetposition_Click()
  FrameMagnet.Visible = Not FrameMagnet.Visible
  If FrameMagnet.Visible Then
     Magnetposition.caption = "Bundle in progress"
  Else
     Magnetposition.caption = "Magnet position"
  End If
End Sub

Private Sub SelScivolo_Click()
    DB473.Bit(62, 3) = Not DB473.Bit(62, 3)
    If DB473.Bit(62, 3) = True Then
        SelScivolo.Picture = ImageList1.ListImages(2).Picture
    Else
        SelScivolo.Picture = ImageList1.ListImages(1).Picture
    End If
End Sub

' aggiornamento dati della pagina corrente

Private Sub TimerLocale_Timer()
   DB420.Refresh
   Me.Update
End Sub


'**************************************************
' Funzione di aggiornamento video e comunicazione
' con il plc da chiamare in background
'**************************************************
Public Sub Update()
    Static Delay As Boolean
    
    Dim i As Integer
    
    '============================================
    On Error Resume Next
    
    Delay = Not Delay
    DisplayTubiAccum(0) = Str(DB420.Word(60))
    DisplayTubiAccum(1) = Str(DB420.Word(62))
    If DB402.Bit(0, 3) Then
        TubesLastRow.Visible = True
        lblLastRow.Visible = True
        If Delay Then
           If lblLastRow.BackColor = &HFFC0C0 Then
              lblLastRow.BackColor = &HF0F0&
           Else
              lblLastRow.BackColor = &HFFC0C0
           End If
        End If
    Else
        TubesLastRow.Visible = False
        lblLastRow.Visible = False
    End If
    
    '============================================
    lblMTtot = LblLung / 1000 * (IIf((BundlesCounterDisplay - 1) <= 0, 0, Val((BundlesCounterDisplay - 1))) * BundleTubesPresetDisplay + BundleTubesCounterDisplay)
    lblTubesTot = (IIf((BundlesCounterDisplay - 1) <= 0, 0, Val((BundlesCounterDisplay - 1))) * BundleTubesPresetDisplay + BundleTubesCounterDisplay)
    lblpeso = Val(DB470.Word(12) / 100 * (IIf((BundlesCounterDisplay - 1) <= 0, 0, Val((BundlesCounterDisplay - 1))) * BundleTubesPresetDisplay + BundleTubesCounterDisplay))

    '============================================
    'controlla se sono cambiati i tubi per fila
    ChangeON = False
    For i = 1 To 50
        If DB470.Word(78 + i * 2) <> MemOld(i) Then ChangeON = True
        MemOld(i) = DB470.Word(78 + i * 2)
    Next i
    'controlla se sono cambiati i dati del tubo
    TubeChange = False
    If OldTube.OldAl <> DB470.Word(8) Or OldTube.OldLa <> DB470.Word(6) Or _
       OldTube.OldSp <> DB470.Word(10) Or ChangeON Or _
       OldTube.OldTi <> DB470.Bit(2, 0) Then TubeChange = True
       
    OldTube.OldAl = DB470.Word(8)
    OldTube.OldLa = DB470.Word(6)
    OldTube.OldSp = DB470.Word(10)
    OldTube.OldTi = DB470.Bit(2, 0)
    '============================================
    '============================================
    
    ' aggiorna dati magneti
    PosX(0) = Format(Conv_UM.Conversione(DB420.DWord(48) / 1000, UM.mm, UM.inch, 3), "###0.000")
    PosX(6) = Format(Conv_UM.Conversione(DB420.DWord(52) / 1000, UM.mm, UM.inch, 3), "###0.000")
    
    ImgAsseMag(0).Visible = Not DBTestIN.Bit(14, 2, PlcIN)
    ImgAsseMag(1).Visible = Not DBTestIN.Bit(14, 3, PlcIN)
    ImgAsseMag(2).Visible = Not DBTestIN.Bit(14, 4, PlcIN)
    ImgAsseMag(3).Visible = Not DBTestIN.Bit(14, 5, PlcIN)
    ImgAsseMag(4).Visible = Not DBTestIN.Bit(14, 6, PlcIN)
    ImgAsseMag(5).Visible = Not DBTestIN.Bit(14, 7, PlcIN)
    ImageMagnete(0).Move ImageMagnete(0).Left, 1380 + Abs(ImgAsseMag(0).Visible) * 760
    ImageMagnete(1).Move ImageMagnete(1).Left, 1380 + Abs(ImgAsseMag(1).Visible) * 760
    ImageMagnete(2).Move ImageMagnete(2).Left, 1380 + Abs(ImgAsseMag(2).Visible) * 760
    ImageMagnete(3).Move ImageMagnete(3).Left, 1380 + Abs(ImgAsseMag(3).Visible) * 760
    ImageMagnete(4).Move ImageMagnete(4).Left, 1380 + Abs(ImgAsseMag(4).Visible) * 760
    ImageMagnete(5).Move ImageMagnete(5).Left, 1380 + Abs(ImgAsseMag(5).Visible) * 760
    For i = 0 To 5
       Magnetizzato(i).Move Magnetizzato(i).Left, ImageMagnete(i).Top + 1700
    Next
    For i = 0 To 5
        If DB420.Bit(56, i) Then
           Enab(i).BackColor = vbGreen
        Else
           Enab(i).BackColor = &H808080
        End If
        If DB420.Bit(58, i) Then
           Magnetizzato(i).BackColor = vbRed
        Else
           Magnetizzato(i).BackColor = vbYellow
        End If
    Next

'    If DB420.Bit(58, 1) Then
'       Magnetizzato(1).BackColor = vbRed
'    Else
'       Magnetizzato(1).BackColor = vbYellow
'    End If
'    If DB420.Bit(58, 2) Then
'       Magnetizzato(2).BackColor = vbRed
'    Else
'       Magnetizzato(2).BackColor = vbYellow
'    End If
'    If DB420.Bit(58, 3) Then
'       Magnetizzato(3).BackColor = vbRed
'    Else
'       Magnetizzato(3).BackColor = vbYellow
'    End If
'    If DB420.Bit(58, 4) Then
'       Magnetizzato(4).BackColor = vbRed
'    Else
'       Magnetizzato(4).BackColor = vbYellow
'    End If
'    If DB420.Bit(58, 5) Then
'       Magnetizzato(5).BackColor = vbRed
'    Else
'       Magnetizzato(5).BackColor = vbYellow
'    End If
    '============================================
    '============================================
   ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    SelScivolo.Visible = Param.GetBit("Par207_PresenzaScivoloEntrata")
    
    ' controlla che siano letti i dati dal plc una prima volta per aggiornare lo stato del pacco
    
    lblbar(2) = PaginaPacco.Ordine_Descrizione
    lblbar(4) = PaginaPacco.Ricetta_Descrizione
    LblLung = Conv_UM.Conversione(DB470.Word(4), UM.mm, UM.inch)
    
    ' =======================  scivolo
    If DB473.Bit(62, 3) = True Then
        SelScivolo.Picture = ImageList1.ListImages(2).Picture
    Else
        SelScivolo.Picture = ImageList1.ListImages(1).Picture
    End If
    '===============================
    If DB402.DatiCambiati Then
        BundlesPresetDisplay.caption = DB402.Word(2)
        If DB402.Bit(0, 5) Or DB402.Bit(0, 6) Then
            BundlesPresetDisplay.Visible = True
            OfLabel(1).Visible = True
        Else
            BundlesPresetDisplay.Visible = False
            OfLabel(1).Visible = False
        End If
        DB402.DatiCambiati = False
        DB402.NumCambiamenti = DB402.NumCambiamenti - 1
     End If
        
    ' aggiorna selettori
    
    If DB473.DatiCambiati Then
        If ControlloMagneti.Occupato = False Then
           ControlloMagneti.value = DB473.Word(64)
           ControlloMagneti.Refresh
        End If
        If DB473.Bit(62, 0) = True Then
            SelettorePale.Picture = ImageList1.ListImages(2).Picture
        Else
            SelettorePale.Picture = ImageList1.ListImages(1).Picture
        End If
        If DB473.Bit(62, 1) = True Then
            SelettoreSagome.Picture = ImageList1.ListImages(2).Picture
        Else
            SelettoreSagome.Picture = ImageList1.ListImages(1).Picture
        End If
        DB473.DatiCambiati = False
        DB473.NumCambiamenti = DB473.NumCambiamenti - 1
    End If

    '=========================== DISEGNO TUBO-PACCO ============================
        ' aConfig = Selezioni varie
        '           bit 0 - peso 1 : unità di misura : 0 = mm   1 = inch
        '           bit 1 - peso 2: tipo di disegno : 0 = pacco 1=tubo
        '           bit 2 - peso 4 disegna spessore: 0 = no 1=si
        '           bit 3 - peso 8 disegna label   : 0 = no 1=si
        ' aTipoTubo => tubo tondo = 1
        ' aTipoPacco => pacco esagono = 1
        ' aTube_Width = larghezza tubo in   mm x 10   o inch * 100
        ' aTube_Height = altezza tubo in    mm x 10   o inch * 100
        ' aTube_Tickness = spessore tubo in mm x 100  o inch * 1000
        ' aCounted  = Tubi presenti nel pacco
        ' Row_01 .... Row_50 = Numero tubi per ogni fila
        
        'TUBO
       
         If nTubi <> DB420.Word(30) Or PaginaAttiva Or DB470.DatiCambiati Or ChangeON Or TubeChange Then
            'numero pacchi fatti
            DisegnoPacco.aCounted = DB420.Word(30)
            nTubi = DB420.Word(30)
            'numero tubi nel pacco
            BundleTubesCounterDisplay.caption = DB420.Word(30)
            BundlesCounterDisplay.caption = DB420.Word(32)
            
            If ControlloTempoAllineamentoTubo.Occupato = False Then
               ControlloTempoAllineamentoTubo.value = DB470.Word(64) / 10#
               ControlloTempoAllineamentoTubo.Refresh
            End If
            BundleTubesPresetDisplay = DB470.Word(22)
            ' dati tubo cambiati
            If TubeChange Or PaginaAttiva Then
                DisegnoTubo.aConfig = Param.GetNumber("Par101_MisureMetriche") + 2 + 4 + 8
                DisegnoTubo.aTube_Height = Conv_UM.Conversione(DB470.Word(8), UM.mm, UM.inch, 2) * 10
                DisegnoTubo.aTube_Width = Conv_UM.Conversione(DB470.Word(6), UM.mm, UM.inch, 2) * 10
                DisegnoTubo.aTube_Tickness = Conv_UM.Conversione(DB470.Word(10), UM.mm, UM.inch, 4) * 100
                If DB470.Bit(2, 0) Then
                    DisegnoTubo.aTipoTubo = Tubo.Tondo
                Else
                    DisegnoTubo.aTipoTubo = Tubo.Quadro
                End If
            End If
            '**************************************** PACCO
            NumProfilo = Abs(DB470.Bit(2, 1)) + 2 * Abs(DB470.Bit(2, 2))
            DisegnoPacco.aConfig = Param.GetNumber("Par101_MisureMetriche") + 4 + 8 + Abs(NumProfilo > 0) * 16
            DisegnoPacco.aTube_Height = Conv_UM.Conversione(DB470.Word(8), UM.mm, UM.inch, 2) * 10
            DisegnoPacco.aTube_Width = Conv_UM.Conversione(DB470.Word(6), UM.mm, UM.inch, 2) * 10
            DisegnoPacco.aTube_Tickness = Conv_UM.Conversione(DB470.Word(10), UM.mm, UM.inch, 4) * 100
            DisegnoPacco.aTipoProfilo = NumProfilo
    
            ' caratteristiche del pacco e del tubo
            
            If DB470.Bit(2, 0) Then
                DisegnoPacco.aTipoTubo = Tubo.Tondo
            Else
                DisegnoPacco.aTipoTubo = Tubo.Quadro
            End If
            
            If DB470.Bit(20, 0) = False Then
               DisegnoPacco.aTipoPacco = Pacco.Quadro
            Else
               DisegnoPacco.aTipoPacco = Pacco.Esagono
            End If
            
            'legge dal plc il numero di tubi per file
            
            For i = 1 To MAX_ROWS
                DisegnoPacco.TubiFila(i) = DB470.Word(i * 2 + 78)
            Next
            
            ' ridisegna il tubo
            
           '  If TubeChange Or PaginaAttiva Then
             DisegnoTubo.Refresh
             
             ' ridisegna il tubo ed il pacco se è cambiato il db 470
             If DB470.NumCambiamenti > 0 Then
                DB470.NumCambiamenti = DB470.NumCambiamenti - 1
             Else
                DB470.NumCambiamenti = 0
             End If
             If DB420.NumCambiamenti > 0 Then
                DB420.NumCambiamenti = DB420.NumCambiamenti - 1
              Else
                DB420.NumCambiamenti = 0
              End If
             DB470.DatiCambiati = False
             DB420.DatiCambiati = False
    End If
    
    If DB470.DatiCambiati Or ChangeON Or TubeChange Then
       DisegnoTubo.Refresh
       DisegnoPacco.Refresh
       DB470.Refresh
    End If
    
End Sub
'

Private Sub Form_Activate()
   PaginaAttiva = True
   RefreshPagina
   PaginaAttiva = False
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
End Sub

Private Sub Form_Load()
   ScritteMultilingua
   TimerLocale.Enabled = False
   WindowState = 2
End Sub

' modifica dati
Private Sub CommandModifica_Click()
    ' legge l'ordine e la ricetta dall'archivio
    ModOrder.IDOrdine = DB470.Word(0)
    OrdersForm.LoadOrderData ModOrder
    Ricetta.IDRicetta = ModOrder.IDRicetta
    OrdersForm.LoadRecipeData Ricetta
    ' sovrascrive con i dati presenti su plc
    UploadDB470 ModOrder, Ricetta
    UploadDB473 Ricetta
    ' modifica
    RecipeModifyForm.StrapEnable = False
    RecipeModifyForm.BundleEnable = True
    RecipeModifyForm.SSTab1.Tab = 0
    RecipeModifyForm.Show vbModal
    If RecipeModifyForm.OK Then
        ' scrive su plc i dati
        DownloadDB470 ModOrder, Ricetta
        PaginaPacco.Ricetta_Descrizione = ModOrder.IDRicetta
        PaginaPacco.Ordine_Descrizione = ModOrder.Descrizione
        DownloadDB473 Ricetta
        ' salva le modifiche
        OrdersForm.SaveOrderData ModOrder
        OrdersForm.SaveRecipeData Ricetta
        OrderChanged = True
    End If
    RefreshPagina
End Sub
Public Sub UploadDB473(RecipeDest As RecipeClass)
   
   RecipeDest.VelMPS = DB473.Word(64)
   
End Sub
' lettura dei dati da DB470
Public Sub UploadDB470(OrderDest As OrderClass, RecipeDest As RecipeClass)
    Dim i As Integer
    On Error Resume Next    ' per eventuali overflow
        OrderDest.IDOrdine = DB470.Word(0)
        If DB470.Bit(2, 0) = True Then
           RecipeDest.TipoTubo = 1
        Else
           RecipeDest.TipoTubo = 2
        End If
        RecipeDest.TuboLunghezza = DB470.Word(4) / 1000#
        RecipeDest.TuboLarghezza = DB470.Word(6) / 10000#
        RecipeDest.TuboAltezza = DB470.Word(8) / 10000#
        RecipeDest.TuboSpessore = DB470.Word(10) / 10000#
        If DB470.Bit(20, 0) = True Then
           RecipeDest.TipoPacco = Pacco.Esagono
        Else
           RecipeDest.TipoPacco = Pacco.Quadro
        End If
        For i = 1 To 50
            RecipeDest.TubiFila(i) = DB470.Word(78 + i * 2)
        Next i
        ' ricalcola gli altri dati (dimensioni e peso pacco)
        RecipeDest.ControlloDatiPacco
   
        If DB402.Bit(0, 5) Then
            OrderDest.ModoCambioOrdine = 2
        Else
            If DB402.Bit(0, 6) Then
                OrderDest.ModoCambioOrdine = 1
            Else
                OrderDest.ModoCambioOrdine = 0
            End If
        End If
        OrderDest.PresetPacchi = DB402.Word(2)
    On Error GoTo 0
End Sub

Private Sub SelettorePale_Click()
   DB473.Bit(62, 0) = Not DB473.Bit(62, 0)
    If DB473.Bit(62, 0) = True Then
        SelettorePale.Picture = ImageList1.ListImages(2).Picture
    Else
        SelettorePale.Picture = ImageList1.ListImages(1).Picture
    End If
End Sub

Public Sub AggiornaDaUpDownControl()
    DB470.Word(64) = ControlloTempoAllineamentoTubo.value * 10
    DB473.Word(64) = ControlloMagneti.value
End Sub

Private Sub SelettoreSagome_Click()
    DB473.Bit(62, 1) = Not DB473.Bit(62, 1)
    If DB473.Bit(62, 1) = True Then
        SelettoreSagome.Picture = ImageList1.ListImages(2).Picture
    Else
        SelettoreSagome.Picture = ImageList1.ListImages(1).Picture
    End If
End Sub

Sub ScritteMultilingua()
'   Label1.caption = Param.Text("ALLINEAMENTO (s)")
   Label3.caption = Param.Text("SAGOME HEX")
   Label2.caption = Param.Text("PALE VELOCI")
   Label8.caption = Param.Text("Vel. magneti (%)")
   CommandModifica.caption = Param.Text("MODIFICA")
   CurrentTubesLabel.caption = Param.Text("Tubi")
   OfLabel(0).caption = Param.Text("di")
   OfLabel(1).caption = Param.Text("di")
   CurrentBundleLabel.caption = Param.Text("Pacco n.")
   lblbar(5) = Param.Text("Bundle page")
   lblbar(1) = Param.Text("Pagina")
   lblbar(3) = Param.Text("Ricette")
   lblbar(0) = Param.Text("ORDER")
   LblScivolo = Param.Text("Scivolo")
   Label01(2) = Param.Text("Scivolo01")
   Label7 = Param.Text("000000079")
   Label9 = Param.Text("000000080")
   Label10 = Param.Text("000000081")
   Label11 = Param.Text("TubiCatene")
End Sub

Private Sub TubesLastRow_Click()
            TOUCHNumericPad.ValoreMin = 0
            TOUCHNumericPad.ValoreMax = 10
            TOUCHNumericPad.Dati = DB420.Word(38) 'Tubi per fare la fila
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                DB402.Word(4) = TOUCHNumericPad.Dati
                TubesLastRow.caption = TOUCHNumericPad.Dati
                DB402.Bit(0, 3) = False
            End If
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
           Unload frmHelp
           Set frmHelp = Nothing
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
                .NomeFile = "Pacco_pagina.htm"
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

Sub RefreshPagina()
    Dim i As Integer
    
    For i = 0 To 5
       ImgAsseMag(i).Visible = False
    Next
    ' refresh della pagina all'attivazione
    FrameMagnet.Visible = False
    Me.Update
    DB473.NumCambiamenti = 0
    DB470.NumCambiamenti = 0
    DB402.NumCambiamenti = 0
    DB420.NumCambiamenti = 0
    ' controllo lunghezza
    ImgLung.Picture = LoadPicture("..\bitmap\TuboLunghezza.gif")
    ' abilitazione temporizzatore locale
    TimerLocale.Enabled = True
    TimerLocale.Interval = 100
  
    PrimaScansione = False
    'Barra1.Pulsante_Click 4
    Barra21.Selezionato = 4
    
    ' pulsanti PP / MPS
    Label8.Visible = Param.GetBit("Par214_MPS")
    ControlloMagneti.Visible = Label8.Visible
    Magnetposition.Visible = Label8.Visible
   
    ' =======================  scivolo
     If Param.GetBit("Par207_PresenzaScivoloEntrata") Then
        LblScivolo.Visible = True
        SelScivolo.Visible = True
        Label01(2).Visible = True
    Else
        LblScivolo.Visible = False
        SelScivolo.Visible = False
        Label01(2).Visible = False
    End If
    
     If DB473.Bit(62, 3) = True Then
        SelScivolo.Picture = ImageList1.ListImages(2).Picture
    Else
        SelScivolo.Picture = ImageList1.ListImages(1).Picture
    End If
    '===============================
    If Param.GetBit("Par215_SelettorePale") Then
       Label2.Visible = True
       Label01(0).Visible = True
       SelettorePale.Visible = True
       ShapeRif.Visible = True
    Else
       ShapeRif.Visible = False
       Label2.Visible = False
       Label01(0).Visible = False
       SelettorePale.Visible = False
    End If
    
    If Param.GetBit("Par216_Selettore controsagome") Then
       Label3.Visible = True
       Label01(1).Visible = True
       SelettoreSagome.Visible = True
       Shape1.Visible = True
    Else
       Label3.Visible = False
       Label01(1).Visible = False
       SelettoreSagome.Visible = False
       Shape1.Visible = False
    End If
    ControlloTempoAllineamentoTubo.Step = 0.1
    ControlloTempoAllineamentoTubo.LimMax = 10
    ControlloTempoAllineamentoTubo.LimMin = 0.1
    ControlloTempoAllineamentoTubo.Decimali = 1
    ControlloTempoAllineamentoTubo.value = DB470.Word(64) / 10#
    ControlloTempoAllineamentoTubo.Refresh
    ControlloMagneti.Step = 10
    ControlloMagneti.LimMax = 100
    ControlloMagneti.LimMin = 20
    ControlloMagneti.Decimali = 0
    ControlloMagneti.value = DB473.Word(64)
    ControlloMagneti.Refresh
    DisegnoPacco.ColoreSfondo = &HC0C0C0
    DisegnoTubo.ColoreSfondo = &HC0C0C0
    RecipeModifyForm.SSTab1.TabEnabled(4) = False
End Sub
