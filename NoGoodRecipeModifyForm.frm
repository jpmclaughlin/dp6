VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form RecipeModifyForm 
   BackColor       =   &H0080C0FF&
   Caption         =   "Recipe modify"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CancelCommand 
      Cancel          =   -1  'True
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2415
      TabIndex        =   8
      Top             =   10020
      Width           =   2295
   End
   Begin VB.CommandButton OkCommand 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   45
      TabIndex        =   7
      Top             =   10020
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   -15
      TabIndex        =   9
      Top             =   45
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   17595
      _Version        =   393216
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Bundle modify"
      TabPicture(0)   =   "RecipeModifyForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePaccoTubo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Strap modify"
      TabPicture(1)   =   "RecipeModifyForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SpecialStrapFrame"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Preset"
      TabPicture(2)   =   "RecipeModifyForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SpecialFrame"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "TubeDimensionFrame"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame TubeDimensionFrame 
         BackColor       =   &H00FFFF80&
         Caption         =   "Dimensioni tubo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74925
         TabIndex        =   145
         Top             =   1635
         Width           =   11655
         Begin VB.Label DiameterLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Diametro"
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
            Left            =   3180
            TabIndex        =   0
            Top             =   495
            Width           =   2775
         End
         Begin VB.Label DimensionLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Dimensioni (mm)"
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
            Left            =   300
            TabIndex        =   1
            Top             =   420
            Width           =   2835
         End
         Begin VB.Label LengthLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Lunghezza"
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
            Left            =   3180
            TabIndex        =   148
            Top             =   1740
            Width           =   2415
         End
         Begin VB.Label XLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
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
            Left            =   7500
            TabIndex        =   147
            Top             =   480
            Width           =   495
         End
         Begin VB.Label ThickLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Spessore"
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
            Left            =   3180
            TabIndex        =   146
            Top             =   1140
            Width           =   2415
         End
      End
      Begin VB.Frame SpecialFrame 
         BackColor       =   &H0080FF80&
         Height          =   6915
         Left            =   -71460
         TabIndex        =   42
         Top             =   2175
         Width           =   8715
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   30
            Left            =   5805
            TabIndex        =   94
            Text            =   "RowEdit(30)"
            Top             =   5460
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   31
            Left            =   5805
            TabIndex        =   93
            Text            =   "RowEdit(31)"
            Top             =   4880
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   32
            Left            =   5805
            TabIndex        =   92
            Text            =   "RowEdit(32)"
            Top             =   4300
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   33
            Left            =   5805
            TabIndex        =   91
            Text            =   "RowEdit(33)"
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   34
            Left            =   5805
            TabIndex        =   90
            Text            =   "RowEdit(34)"
            Top             =   3140
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   35
            Left            =   5805
            TabIndex        =   89
            Text            =   "RowEdit(35)"
            Top             =   2560
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   36
            Left            =   5805
            TabIndex        =   88
            Text            =   "RowEdit(36)"
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   37
            Left            =   5805
            TabIndex        =   87
            Text            =   "RowEdit(37)"
            Top             =   1400
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   38
            Left            =   5805
            TabIndex        =   86
            Text            =   "RowEdit(38)"
            Top             =   820
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   39
            Left            =   5805
            TabIndex        =   85
            Text            =   "RowEdit(39)"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   47
            Left            =   7560
            TabIndex        =   84
            Text            =   "RowEdit(39)"
            Top             =   1400
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   49
            Left            =   7560
            TabIndex        =   83
            Text            =   "RowEdit(39)"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   48
            Left            =   7560
            TabIndex        =   82
            Text            =   "RowEdit(39)"
            Top             =   820
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   44
            Left            =   7560
            TabIndex        =   81
            Text            =   "RowEdit(39)"
            Top             =   3140
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   46
            Left            =   7560
            TabIndex        =   80
            Text            =   "RowEdit(39)"
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   45
            Left            =   7560
            TabIndex        =   79
            Text            =   "RowEdit(39)"
            Top             =   2560
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   43
            Left            =   7560
            TabIndex        =   78
            Text            =   "RowEdit(39)"
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   42
            Left            =   7560
            TabIndex        =   77
            Text            =   "RowEdit(39)"
            Top             =   4300
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   41
            Left            =   7560
            TabIndex        =   76
            Text            =   "RowEdit(39)"
            Top             =   4880
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   40
            Left            =   7560
            TabIndex        =   75
            Text            =   "RowEdit(39)"
            Top             =   5460
            Width           =   975
         End
         Begin VB.CommandButton SaveSpecialBundleCommand 
            Caption         =   "Visualizza disegno pacco"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            TabIndex        =   74
            Top             =   6120
            Width           =   4035
         End
         Begin VB.CommandButton CancelSpecialCommand 
            Caption         =   "Annulla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4500
            TabIndex        =   73
            Top             =   6120
            Width           =   4035
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   29
            Left            =   4050
            TabIndex        =   72
            Text            =   "RowEdit(29)"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   28
            Left            =   4050
            TabIndex        =   71
            Text            =   "RowEdit(28)"
            Top             =   820
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   27
            Left            =   4050
            TabIndex        =   70
            Text            =   "RowEdit(27)"
            Top             =   1400
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   26
            Left            =   4050
            TabIndex        =   69
            Text            =   "RowEdit(26)"
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   25
            Left            =   4050
            TabIndex        =   68
            Text            =   "RowEdit(25)"
            Top             =   2560
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   24
            Left            =   4050
            TabIndex        =   67
            Text            =   "RowEdit(24)"
            Top             =   3140
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   23
            Left            =   4050
            TabIndex        =   66
            Text            =   "RowEdit(23)"
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   22
            Left            =   4050
            TabIndex        =   65
            Text            =   "RowEdit(22)"
            Top             =   4300
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   21
            Left            =   4050
            TabIndex        =   64
            Text            =   "RowEdit(21)"
            Top             =   4880
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   20
            Left            =   4050
            TabIndex        =   63
            Text            =   "RowEdit(20)"
            Top             =   5460
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   19
            Left            =   2220
            TabIndex        =   62
            Text            =   "RowEdit(19)"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   18
            Left            =   2220
            TabIndex        =   61
            Text            =   "RowEdit(18)"
            Top             =   820
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   17
            Left            =   2220
            TabIndex        =   60
            Text            =   "RowEdit(17)"
            Top             =   1400
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   16
            Left            =   2220
            TabIndex        =   59
            Text            =   "RowEdit(16)"
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   15
            Left            =   2220
            TabIndex        =   58
            Text            =   "RowEdit(15)"
            Top             =   2560
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   14
            Left            =   2220
            TabIndex        =   57
            Text            =   "RowEdit(14)"
            Top             =   3140
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   13
            Left            =   2220
            TabIndex        =   56
            Text            =   "RowEdit(13)"
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   12
            Left            =   2220
            TabIndex        =   55
            Text            =   "RowEdit(12)"
            Top             =   4300
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   10
            Left            =   2220
            TabIndex        =   54
            Text            =   "RowEdit(10)"
            Top             =   5460
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   11
            Left            =   2220
            TabIndex        =   53
            Text            =   "RowEdit(11)"
            Top             =   4880
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   540
            TabIndex        =   52
            Text            =   "RowEdit(0)"
            Top             =   5460
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   540
            TabIndex        =   51
            Text            =   "RowEdit(1)"
            Top             =   4875
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   2
            Left            =   540
            TabIndex        =   50
            Text            =   "RowEdit(2)"
            Top             =   4305
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   3
            Left            =   540
            TabIndex        =   49
            Text            =   "RowEdit(3)"
            Top             =   3720
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   4
            Left            =   540
            TabIndex        =   48
            Text            =   "RowEdit(4)"
            Top             =   3135
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   5
            Left            =   540
            TabIndex        =   47
            Text            =   "RowEdit(5)"
            Top             =   2565
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   6
            Left            =   540
            TabIndex        =   46
            Text            =   "RowEdit(6)"
            Top             =   1980
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   7
            Left            =   540
            TabIndex        =   45
            Text            =   "RowEdit(7)"
            Top             =   1395
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   8
            Left            =   540
            TabIndex        =   44
            Text            =   "RowEdit(8)"
            Top             =   825
            Width           =   975
         End
         Begin VB.TextBox RowEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   9
            Left            =   540
            TabIndex        =   43
            Text            =   "RowEdit(9)"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "50"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   49
            Left            =   7140
            TabIndex        =   144
            Top             =   360
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   48
            Left            =   7140
            TabIndex        =   143
            Top             =   960
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "48"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   47
            Left            =   7140
            TabIndex        =   142
            Top             =   1520
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "47"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   46
            Left            =   7140
            TabIndex        =   141
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "46"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   45
            Left            =   7140
            TabIndex        =   140
            Top             =   2680
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "45"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   44
            Left            =   7140
            TabIndex        =   139
            Top             =   3260
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "44"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   43
            Left            =   7140
            TabIndex        =   138
            Top             =   3840
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "43"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   42
            Left            =   7140
            TabIndex        =   137
            Top             =   4420
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "42"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   41
            Left            =   7140
            TabIndex        =   136
            Top             =   5000
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "41"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   40
            Left            =   7140
            TabIndex        =   135
            Top             =   5580
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   29
            Left            =   3570
            TabIndex        =   134
            Top             =   360
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "29"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   3570
            TabIndex        =   133
            Top             =   940
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "28"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   27
            Left            =   3570
            TabIndex        =   132
            Top             =   1520
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   26
            Left            =   3570
            TabIndex        =   131
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "26"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   3570
            TabIndex        =   130
            Top             =   2680
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   3570
            TabIndex        =   129
            Top             =   3260
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   23
            Left            =   3570
            TabIndex        =   128
            Top             =   3840
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   22
            Left            =   3570
            TabIndex        =   127
            Top             =   4420
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   21
            Left            =   3570
            TabIndex        =   126
            Top             =   5000
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   20
            Left            =   3570
            TabIndex        =   125
            Top             =   5580
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "40"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   39
            Left            =   5355
            TabIndex        =   124
            Top             =   360
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "39"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   38
            Left            =   5355
            TabIndex        =   123
            Top             =   940
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "38"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   37
            Left            =   5355
            TabIndex        =   122
            Top             =   1520
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   36
            Left            =   5355
            TabIndex        =   121
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "36"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   5355
            TabIndex        =   120
            Top             =   2680
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "35"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   34
            Left            =   5355
            TabIndex        =   119
            Top             =   3260
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "34"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   33
            Left            =   5355
            TabIndex        =   118
            Top             =   3840
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   32
            Left            =   5355
            TabIndex        =   117
            Top             =   4420
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   31
            Left            =   5355
            TabIndex        =   116
            Top             =   5000
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   30
            Left            =   5355
            TabIndex        =   115
            Top             =   5580
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   19
            Left            =   1740
            TabIndex        =   114
            Top             =   360
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   18
            Left            =   1740
            TabIndex        =   113
            Top             =   940
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   17
            Left            =   1740
            TabIndex        =   112
            Top             =   1520
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   1740
            TabIndex        =   111
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   1740
            TabIndex        =   110
            Top             =   2680
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   1740
            TabIndex        =   109
            Top             =   3260
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   13
            Left            =   1740
            TabIndex        =   108
            Top             =   3840
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   1740
            TabIndex        =   107
            Top             =   4420
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   1740
            TabIndex        =   106
            Top             =   5000
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   1785
            TabIndex        =   105
            Top             =   5580
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   104
            Top             =   4944
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   103
            Top             =   5580
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   60
            TabIndex        =   102
            Top             =   360
            Width           =   375
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   101
            Top             =   933
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   100
            Top             =   1506
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   99
            Top             =   2079
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   98
            Top             =   2652
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   97
            Top             =   3225
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   96
            Top             =   3798
            Width           =   255
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   95
            Top             =   4371
            Width           =   255
         End
      End
      Begin VB.Frame SpecialStrapFrame 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9075
         Left            =   -74865
         TabIndex        =   13
         Top             =   1245
         Width           =   14655
         Begin VB.Frame StrapDrawFrame 
            BackColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   -45
            TabIndex        =   37
            Top             =   6390
            Width           =   14655
            Begin VB.CommandButton SpecialStrapCommand 
               Caption         =   "Modifica reggie"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   2460
               TabIndex        =   39
               Top             =   1440
               Width           =   3975
            End
            Begin VB.CommandButton DeleteSpecialStrapCommand 
               Caption         =   "Cancella configurazione reggie"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               Left            =   7800
               TabIndex        =   38
               Top             =   1440
               Width           =   3975
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   1
               X1              =   1620
               X2              =   1620
               Y1              =   255
               Y2              =   1095
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   2
               X1              =   1425
               X2              =   1425
               Y1              =   255
               Y2              =   1095
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   10
               X1              =   3330
               X2              =   3330
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   9
               X1              =   3105
               X2              =   3105
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   8
               X1              =   2895
               X2              =   2895
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   7
               X1              =   2685
               X2              =   2685
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   6
               X1              =   2475
               X2              =   2475
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   5
               X1              =   2265
               X2              =   2265
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   4
               X1              =   2040
               X2              =   2040
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   3
               X1              =   1830
               X2              =   1830
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   0
               X1              =   1155
               X2              =   1155
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Line StrapLine 
               BorderWidth     =   6
               Index           =   11
               X1              =   3540
               X2              =   3540
               Y1              =   240
               Y2              =   1080
            End
            Begin VB.Shape BundleShape 
               BackColor       =   &H80000003&
               BackStyle       =   1  'Opaque
               FillStyle       =   2  'Horizontal Line
               Height          =   855
               Left            =   3540
               Top             =   240
               Width           =   8655
            End
         End
         Begin VB.TextBox PosizioneDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   10740
            MaxLength       =   5
            TabIndex        =   32
            Text            =   "0"
            Top             =   3180
            Width           =   2595
         End
         Begin VB.CommandButton IncStrapCommand 
            Height          =   855
            Left            =   5340
            Picture         =   "RecipeModifyForm.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2340
            Width           =   2295
         End
         Begin VB.CommandButton DecStrapCommand 
            Height          =   855
            Left            =   5340
            Picture         =   "RecipeModifyForm.frx":07CD
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4200
            Width           =   2295
         End
         Begin VB.TextBox StrapsDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   5520
            MaxLength       =   2
            TabIndex        =   29
            Text            =   "0"
            Top             =   3360
            Width           =   1935
         End
         Begin VB.TextBox FirstDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   5520
            MaxLength       =   4
            TabIndex        =   28
            Text            =   "0"
            Top             =   1020
            Width           =   1935
         End
         Begin VB.CommandButton CancelSpecialStrapCommand 
            Caption         =   "Annulla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   8100
            TabIndex        =   27
            Top             =   5880
            Width           =   3975
         End
         Begin VB.CommandButton SaveSpecialStrapCommand 
            Caption         =   "Visualizza disegno reggie"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   2640
            TabIndex        =   26
            Top             =   5820
            Width           =   3975
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1560
            TabIndex        =   25
            Text            =   "0"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2640
            TabIndex        =   24
            Text            =   "1"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   3720
            TabIndex        =   23
            Text            =   "2"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   4800
            TabIndex        =   22
            Text            =   "3"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   5880
            TabIndex        =   21
            Text            =   "4"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   6960
            TabIndex        =   20
            Text            =   "5"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   6
            Left            =   8040
            TabIndex        =   19
            Text            =   "6"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   9120
            TabIndex        =   18
            Text            =   "7"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   10200
            TabIndex        =   17
            Text            =   "8888"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   11280
            TabIndex        =   16
            Text            =   "9"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   10
            Left            =   12360
            TabIndex        =   15
            Text            =   "10"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.TextBox QuoteEdit 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   13440
            TabIndex        =   14
            Text            =   "11"
            Top             =   5160
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Position (mm)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   9420
            TabIndex        =   36
            Top             =   2280
            Width           =   4815
         End
         Begin VB.Label QuoteLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Quote :"
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
            Left            =   180
            TabIndex        =   35
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label StrapsLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Numero di reggie"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   300
            TabIndex        =   34
            Top             =   3420
            Width           =   4635
         End
         Begin VB.Label FirstLabel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Posizione prima reggia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   660
            TabIndex        =   33
            Top             =   1080
            Width           =   4695
         End
      End
      Begin VB.Frame FramePaccoTubo 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6915
         Left            =   225
         TabIndex        =   10
         Top             =   2475
         Width           =   14430
         Begin VB.TextBox ThickDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   5445
            MaxLength       =   6
            TabIndex        =   2
            Text            =   "0"
            Top             =   825
            Width           =   1215
         End
         Begin VB.TextBox HeightDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   12300
            MaxLength       =   6
            TabIndex        =   3
            Text            =   "0"
            Top             =   2865
            Width           =   1215
         End
         Begin VB.TextBox WidthDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9225
            MaxLength       =   6
            TabIndex        =   4
            Text            =   "0"
            Top             =   5475
            Width           =   1215
         End
         Begin VB.TextBox LengthDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5415
            MaxLength       =   6
            TabIndex        =   5
            Text            =   "0"
            Top             =   5985
            Width           =   1650
         End
         Begin VB.OptionButton SquareBundleOption 
            DownPicture     =   "RecipeModifyForm.frx":0F2F
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Left            =   420
            Picture         =   "RecipeModifyForm.frx":16DF
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2640
            Width           =   1875
         End
         Begin VB.OptionButton HexOption 
            DownPicture     =   "RecipeModifyForm.frx":1AAC
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Left            =   420
            Picture         =   "RecipeModifyForm.frx":2661
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   660
            Width           =   1875
         End
         Begin dp6.Pacco6 PaccoModRicetta 
            Height          =   4710
            Left            =   7920
            TabIndex        =   40
            Top             =   870
            Width           =   4485
            _extentx        =   7911
            _extenty        =   8308
         End
         Begin dp6.Pacco6 TuboModRicetta 
            Height          =   3090
            Left            =   3720
            TabIndex        =   41
            Top             =   1950
            Width           =   3000
            _extentx        =   5292
            _extenty        =   5450
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Lunghezza"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1710
            TabIndex        =   6
            Top             =   6120
            Width           =   2340
         End
      End
   End
End
Attribute VB_Name = "RecipeModifyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
'*************************************************
' Costanti e variabili per disegno pacco
'**************************************************
'            Private Const LineWidthPixel As Integer = 4 ' spessore tubi disegnati in pixel
'            Private Const LineWidthTwip As Integer = 50 ' spessore tubi disegnati in twip
'            Private Const ArrowWidth   As Integer = 40
'            Private Const ArrowLength  As Integer = 100

'            ' dati per disegno non scalati (m)
'            Private RQuoteX(5) As Double        'punti per disegno frecce di quotatura
'            Private RQuoteY(5) As Double        'punti per disegno frecce di quotatura
'            Private RTubeX(MAX_TUBES) As Double 'punti per disegno tubi pacco
'            Private RTubeY(MAX_TUBES) As Double 'punti per disegno tubi pacco

'            ' dati per disegno scalati  (twips)
'            Private PQuoteX(5) As Long        'punti per disegno frecce di quotatura
'            Private PQuoteY(5) As Long        'punti per disegno frecce di quotatura
'            Private PTubeX(MAX_TUBES) As Long 'punti per disegno tubi pacco
'            Private PTubeY(MAX_TUBES) As Long 'punti per disegno tubi pacco
'            Private PTubeHeight As Long       'altezza tubo
'            Private PTubeWidth As Long        'larghezza tubo

'*************************************************
' Fine variabili per disegno pacco
'**************************************************
'*************************************************
' Costanti e variabili per disegno reggie
'**************************************************
Dim TwipOffset As Long    ' twips
Dim TwipCoeff As Double   ' twip/m
'*************************************************
' Fine variabili per disegno reggie
'**************************************************

'*************************************************
' Costanti e variabili per gestione input dati
'**************************************************
' Ok è public perchè può essere esaminata per verificare
' se i dati sono stati confermati
Public Ok As Boolean
' flag protezione modifica  dati
Private FromOperator As Boolean
' memoria modifica pacco speciale in corso
Private SpecialBundleFlag As Boolean
Private SpecialStrapFlag As Boolean
' finestra visualizzazione ricette di reggiatura (+ o - 2.5 cm = 5cm )
Private Const LengthWindow As Double = 0.025
'*************************************************
' Fine costanti e variabili per gestione input dati
'**************************************************

'Private Sub TubesDisplay_Click()
'    TOUCHNumericPad.Dati = TubesDisplay.Text
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        TubesDisplay.Text = TOUCHNumericPad.Dati
'    End If
'End Sub

'Private Sub BundlesEdit_Click()
'    TOUCHNumericPad.Dati = OrderBundles.Text
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        OrderBundles.Text = TOUCHNumericPad.Dati
'    End If
'End Sub

Private Sub WidthDisplay_CLICK()
    TOUCHNumericPad.Dati = WidthDisplay.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        WidthDisplay.Text = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub HeightDisplay_CLICK()
    TOUCHNumericPad.Dati = HeightDisplay.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        HeightDisplay.Text = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub ThickDisplay_CLICK()
    TOUCHNumericPad.Dati = ThickDisplay.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        ThickDisplay.Text = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub LengthDisplay_CLICK()
    TOUCHNumericPad.Dati = LengthDisplay.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        LengthDisplay.Text = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub RowEdit_Click(Index As Integer)
    Select Case Index
        Case 0
            TOUCHNumericPad.Dati = RowEdit(0).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(0).Text = TOUCHNumericPad.Dati
            End If
        Case 1
            TOUCHNumericPad.Dati = RowEdit(1).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(1).Text = TOUCHNumericPad.Dati
            End If
        Case 2
            TOUCHNumericPad.Dati = RowEdit(2).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(2).Text = TOUCHNumericPad.Dati
            End If
        Case 3
            TOUCHNumericPad.Dati = RowEdit(3).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(3).Text = TOUCHNumericPad.Dati
            End If
        Case 4
            TOUCHNumericPad.Dati = RowEdit(4).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(4).Text = TOUCHNumericPad.Dati
            End If
        Case 5
            TOUCHNumericPad.Dati = RowEdit(5).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(5).Text = TOUCHNumericPad.Dati
            End If
        Case 6
            TOUCHNumericPad.Dati = RowEdit(6).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(6).Text = TOUCHNumericPad.Dati
            End If
        Case 7
            TOUCHNumericPad.Dati = RowEdit(7).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(7).Text = TOUCHNumericPad.Dati
            End If
        Case 8
            TOUCHNumericPad.Dati = RowEdit(8).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(8).Text = TOUCHNumericPad.Dati
            End If
        Case 9
            TOUCHNumericPad.Dati = RowEdit(9).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(9).Text = TOUCHNumericPad.Dati
            End If
        Case 10
            TOUCHNumericPad.Dati = RowEdit(10).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(10).Text = TOUCHNumericPad.Dati
            End If
        Case 11
            TOUCHNumericPad.Dati = RowEdit(11).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(11).Text = TOUCHNumericPad.Dati
            End If
        Case 12
            TOUCHNumericPad.Dati = RowEdit(12).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(12).Text = TOUCHNumericPad.Dati
            End If
        Case 13
            TOUCHNumericPad.Dati = RowEdit(13).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(13).Text = TOUCHNumericPad.Dati
            End If
        Case 14
            TOUCHNumericPad.Dati = RowEdit(14).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(14).Text = TOUCHNumericPad.Dati
            End If
        Case 15
            TOUCHNumericPad.Dati = RowEdit(15).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(15).Text = TOUCHNumericPad.Dati
            End If
        Case 16
            TOUCHNumericPad.Dati = RowEdit(16).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(16).Text = TOUCHNumericPad.Dati
            End If
        Case 17
            TOUCHNumericPad.Dati = RowEdit(17).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(17).Text = TOUCHNumericPad.Dati
            End If
        Case 18
            TOUCHNumericPad.Dati = RowEdit(18).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(18).Text = TOUCHNumericPad.Dati
            End If
        Case 19
            TOUCHNumericPad.Dati = RowEdit(19).Text
            TOUCHNumericPad.Show vbModal
            If TOUCHNumericPad.DatiConfermati Then
                RowEdit(19).Text = TOUCHNumericPad.Dati
            End If
    End Select
End Sub

'Private Sub PrintEdit_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            TOUCHKeyBoard.Dati = PrintEdit(0).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(0).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 1
'            TOUCHKeyBoard.Dati = PrintEdit(1).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(1).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 2
'            TOUCHKeyBoard.Dati = PrintEdit(2).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(2).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 3
'            TOUCHKeyBoard.Dati = PrintEdit(3).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(3).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 4
'            TOUCHKeyBoard.Dati = PrintEdit(4).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(4).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 5
'            TOUCHKeyBoard.Dati = PrintEdit(5).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(5).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 6
'            TOUCHKeyBoard.Dati = PrintEdit(6).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(6).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 7
'            TOUCHKeyBoard.Dati = PrintEdit(7).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(7).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 8
'            TOUCHKeyBoard.Dati = PrintEdit(8).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(8).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 9
'            TOUCHKeyBoard.Dati = PrintEdit(9).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(9).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 10
'            TOUCHKeyBoard.Dati = PrintEdit(10).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(10).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 11
'            TOUCHKeyBoard.Dati = PrintEdit(11).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(11).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 12
'            TOUCHKeyBoard.Dati = PrintEdit(12).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(12).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 13
'            TOUCHKeyBoard.Dati = PrintEdit(13).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(13).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 14
'            TOUCHKeyBoard.Dati = PrintEdit(14).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(14).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 15
'            TOUCHKeyBoard.Dati = PrintEdit(15).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(15).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 16
'            TOUCHKeyBoard.Dati = PrintEdit(16).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(16).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 17
'            TOUCHKeyBoard.Dati = PrintEdit(17).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(17).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 18
'            TOUCHKeyBoard.Dati = PrintEdit(18).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(18).Text = TOUCHKeyBoard.Dati
'            End If
'        Case 19
'            TOUCHKeyBoard.Dati = PrintEdit(19).Text
'            TOUCHKeyBoard.Show vbModal
'            If TOUCHKeyBoard.DatiConfermati Then
'                PrintEdit(19).Text = TOUCHKeyBoard.Dati
'            End If
'    End Select
'End Sub

'===============================================================================
'--------------------- INIZIO FUNZIONI PAGINA
'===============================================================================
Private Sub Form_Load()
    Ok = False
    
    SSTab1.TabCaption(0) = Param.Text("Order")
    SSTab1.TabCaption(1) = Param.Text("Bundle")
    SSTab1.TabCaption(2) = Param.Text("Straps")
    
    If Param.Bit("PrinterEnable") Then
        SSTab1.TabEnabled(3) = True
        SSTab1.TabCaption(3) = Param.Text("Ticket")
    Else
        SSTab1.TabEnabled(3) = False
        SSTab1.TabCaption(3) = ""
    End If
    
'    If Param.Bit("recipeEnable") Then
'        FrameRecipe.Visible = True
'    Else
'        FrameRecipe.Visible = False
'    End If
    

    OkCommand.Caption = Param.Text("Ok")
    CancelCommand.Caption = Param.Text("Cancel")
    
'----------------- DISEGNO PACCO-TUBO  ----------------
' aConfig = Selezioni varie
'           bit 0 (1) : unità di misura : 0 = mm   1 = inch
'           bit 1 (2) : tipo di pacco   : 0 = quadro 1 = hex
'           bit 2 (4) : tipo di tubo    : 0 = quadro 1 = tondo
'           bit 3 (8) : tipo di disegno : 0 = pacco 1=tubo
'           bit 4 (16): disegna spessore: 0 = no 1=si
'           bit 5 (32): disegna label   : 0 = no 1=si
' aTube_Width = larghezza tubo in   mm x 10   o inch * 100
' aTube_Height = altezza tubo in    mm x 10   o inch * 100
' aTube_Tickness = spessore tubo in mm x 100  o inch * 1000
' aCounted  = Tubi presenti nel pacco
' Row_01 .... Row_50 = Numero tubi per ogni fila
    
    Dim Tondo As Integer
    If ModRecipe.TipoTubo = 1 Then
        Tondo = 1
    Else
        Tondo = 0
    End If
    
    Dim Hex As Integer
    If ModRecipe.TipoPacco = 1 Then
        Hex = 1
    Else
        Hex = 0
    End If
        
        
'    'TUBO
'    TubeOrderMod.aConfig = Abs(Hex * 2) + Abs(Tondo * 4) + 8 + 16 + 32
'    TubeOrderMod.Item.aTube_Height = ModOrder.Recipe.TuboAltezza * 1000
'    TubeOrderMod.Item.aTube_Width = TuboLarghezza * 1000
'    TubeOrderMod.Item.aTube_Tickness = TuboSpessore * 10000
'
'    'PACCO
'    BundleOrderMod.Item.aConfig = Abs(Hex * 2) + Abs(Tondo * 4) + 0 + 16 + 32
'    BundleOrderMod.Item.aConfig = Abs(Hex * 2) + Abs(Tondo * 4) + 8 + 16 + 32
'    BundleOrderMod.Item.aTube_Height = TuboAltezza * 1000
'    BundleOrderMod.Item.aTube_Width = TuboLarghezza * 1000
'    BundleOrderMod.Item.aTube_Tickness = TuboSpessore * 10000
'    BundleOrderMod.Item.aCounted = 0
'    BundleOrderMod.Item.Row_01 = 1
    
    
    
    
    
'    TubeDimensionFrame.Caption = Param.Text("TubeDimension")
'    TubeShapeFrame.Caption = Param.Text("TubeShape")
'    BundleSelectLabel.Caption = Param.Text("BundleSelect")
'
'    SpecialBudleCommand.Caption = Param.Text("SpecialModify")
'    SaveSpecialBundleCommand.Caption = Param.Text("SaveSpecial") 'visualizza disegno pacco
'    CancelSpecialCommand.Caption = Param.Text("CancelModify")
'    DeleteSpecialCommand.Caption = Param.Text("DeleteSpecial") 'cancella pacco speciale
'
'    SpecialStrapCommand.Caption = Param.Text("SpecialModify")
'    DeleteSpecialStrapCommand.Caption = Param.Text("DeleteSpecial") ' cancella reggie speciali
'    SaveSpecialStrapCommand.Caption = Param.Text("SaveSpecial") 'visualizza disegno reggie
'    CancelSpecialStrapCommand.Caption = Param.Text("CancelModify")
'
'    QuoteLabel.Caption = Param.Text("Quotes")
'
'    ' selezione tipo pacchi e tubi
'    RoundTubeOption.Visible = Param.Bit("RoundTube")
'    SquareTubeOption.Visible = Param.Bit("SquareTube")
'    RectTubeOption.Visible = Param.Bit("SquareTube")
'    HexOption.Visible = Param.Bit("HexBundle")
'    SquareBundleOption.Visible = Param.Bit("SquareBundle")
'
'    ' Titoli tabbed
'    SSTab1.TabCaption(0) = Param.Text("Order")
'    SSTab1.TabCaption(1) = Param.Text("Bundle")
'    If Param.Bit("StrapEnable") Then
'        SSTab1.TabCaption(2) = Param.Text("Straps")
'    Else
'        SSTab1.TabCaption(2) = "-----"
'    End If
'
'    'Label tubo
'    TubeShapeFrame.Caption = Param.Text("TubeShape")
'    DimensionLabel.Caption = Param.Text("TubeDimension") & Unit.mmString
'    DiameterLabel.Caption = Param.Text("Diameter") & Unit.mmString
'    ThickLabel.Caption = Param.Text("Thickness") & Unit.mmString
'    LengthLabel.Caption = Param.Text("TubeLength") & Unit.mmString
'
'    ' Label pacco
'    TubesLabel.Caption = Param.Text("Tubes")
'    BundleShapeFrame.Caption = Param.Text("BundleShape")
'    BundleWeightLabel.Caption = Param.Text("BundleWeight") & Unit.KgString
'
'    ' Label reggie
'    StrapsLabel.Caption = Param.Text("NumOfStraps")
'    FirstLabel.Caption = Param.Text("FirstStrap") & Unit.mmString
'
'End Sub
'
'Private Sub Form_Activate()
'    Ok = False
'    SpecialBundleFlag = False
'    SpecialStrapFlag = False
'End Sub
'
'' funzione da richiamare prima della Form.show per caricare i dati
'Public Sub GetData(Source As OrderClass, OnLineModify As Boolean, TabPosition As Integer)
'
'    SpecialBundleFlag = False
'    SpecialStrapFlag = False
'    ' aggiorna immagine tubi e pacco
'    BundleTabUpdate
'    TubeDimensioDisplayUpdate
'    ' aggiorna immagine reggie
'    StrapTabUpdate
'    ' aggiorna immagine etichetta e quantità programmata
'    LabelTabUpdate
'
'
'    'consenso modifica forma tubo e pacco
'    If OnLineModify Then
'        BundleShapeFrame.Visible = False
'        TubeShapeFrame.Visible = False
'    Else
'        BundleShapeFrame.Visible = True
'        TubeShapeFrame.Visible = True
'    End If
'
'    SSTab1.Tab = TabPosition
'
'    Ok = False
'
'
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Ok = False
'End Sub
'
''***********************************************************
'' Funzioni di risposta ai comandi dell'operatore ok e cancel
''***********************************************************
'Private Sub CancelCommand_Click()
'    Ok = False
'    Me.Hide
End Sub

' caricamento su oggetto order dei dati impostati a video
Private Function CheckData() As Boolean
'    ' variabili temporanee
'    'Dim RetryFlag As Boolean
'    Dim ErrorFlag As Boolean
'
'    ' chiude eventuali ricette speciali in fase di editazione
'    If SpecialBundleFlag Then SaveSpecialBundleCommand_Click
'    If SpecialStrapFlag Then SaveSpecialStrapCommand_Click
'
'    ' verifiche se le dimensioni del tubo sono accettabili
'    'RetryFlag = False
'    ErrorFlag = False
'    If Order.Tube.Length > Param.Number("MaxTubeLength") Then ErrorFlag = True
'    If Order.Tube.Length < Param.Number("MinTubeLength") Then ErrorFlag = True
'    If Order.Tube.Height > Param.Number("MaxTubeHeight") Then ErrorFlag = True
'    If Order.Tube.Height < Param.Number("MinTubeHeight") Then ErrorFlag = True
'    If Order.Tube.Width > Param.Number("MaxTubeWidth") Then ErrorFlag = True
'    If Order.Tube.Width < Param.Number("MinTubeWidth") Then ErrorFlag = True
'    If Order.Tube.Thickness > Param.Number("MaxWallThickness") Then ErrorFlag = True
'    If Order.Tube.Thickness < Param.Number("MinWallThickness") Then ErrorFlag = True
'    If ErrorFlag = False Then
'        ' verifica se le dimensioni del pacco sono accettabili
'        If (Order.Tube.Weight * Order.Bundle.Tubes) > Param.Number("MaxBundleWeight") Then ErrorFlag = True
'        If Order.Bundle.Hex Then
'            If Order.Bundle.Base(Order.Tube.Width) > Param.Number("MaxBundleBase") Then ErrorFlag = True
'            If Order.Bundle.Base(Order.Tube.Width) < Param.Number("MinBundleBase") Then ErrorFlag = True
'            If Order.Bundle.Side(Order.Tube.Height) > Param.Number("MaxBundleSide") Then ErrorFlag = True
'            If Order.Bundle.Side(Order.Tube.Height) < Param.Number("MinBundleSide") Then ErrorFlag = True
'        Else
'            If Order.Bundle.Height(Order.Tube.Height) > Param.Number("MaxBundleHeight") Then ErrorFlag = True
'            If Order.Bundle.Height(Order.Tube.Height) < Param.Number("MinBundleHeight") Then ErrorFlag = True
'            If Order.Bundle.Width(Order.Tube.Width) > Param.Number("MaxBundleWidth") Then ErrorFlag = True
'            If Order.Bundle.Width(Order.Tube.Width) < Param.Number("MinBundleWidth") Then ErrorFlag = True
'        End If
'        If ErrorFlag Then
'            MsgBox Param.Text("BundleOutOfLimits"), vbOKOnly, Param.Text("DataError")
'        End If
'    Else
'        MsgBox Param.Text("TubeOutOfLimits"), vbOKOnly, Param.Text("DataError")
'    End If
'    ' set flag dati confermati e chiusura finestra
'    If ErrorFlag = False Then
'        ' salva ricetta tubo (pacco e reggie sono già salvati
'        ' in quanto si tratta solo di selezionare i puntatori)
'        Order.Tube.PrevBundleId = Order.Bundle.BundleId
'        Order.Tube.PrevStrapId = Order.Strap.StrapId
'        OrdersForm.SaveTubeData Order.Tube
'
'        ' Aggiunta 09.08.2000
'        ' dopo la aggiunta del campo "posizione" nelle regge bisogna salvare
'        ' eventuali modifiche a questo campo
'        OrdersForm.SaveStrapData Order.Strap
'        ' Fine aggiunta 09.08.2000
'
'        ' recupera i nuovi dati cartellino , quantità e ricetta
'        LabelTabUpload
'        ' salva la ricetta utilizzata nell'ordine corrente
'        SaveRecipeData
'
'        CheckData = True
'    Else
'        CheckData = False
'    End If
End Function


Private Sub OkCommand_Click()
    ' set flag dati confermati e chiusura finestra
    If CheckData Then
        ' chiude il form
        Ok = True
        Me.Hide
    Else
        Ok = False
    End If
End Sub


Private Sub SaveRecipeData()
'    On Error Resume Next
'    RecipeData.Recordset.FindFirst ("Item = '" & Order.Item & "'")
'    If RecipeData.Recordset.NoMatch Then
'        ' crea nuova ricetta
'        RecipeData.Recordset.AddNew
'        RecipeData.Recordset.Fields("Item") = Order.Item
'    Else
'        ' modifica ricetta esistente
'        RecipeData.Recordset.Edit
'    End If
'    RecipeData.Recordset.Fields("TubeId") = Order.Tube.TubeId
'    RecipeData.Recordset.Fields("BundleId") = Order.Bundle.BundleId
'    RecipeData.Recordset.Fields("StrapId") = Order.Strap.StrapId
'    RecipeData.Recordset.Update
'    RecipeData.Refresh
End Sub

Private Sub LoadRecipeData()
'    Dim PrevId As String
'    RecipeData.Recordset.FindFirst ("Item = '" & ItemDisplay.Text & "'")
'    If RecipeData.Recordset.NoMatch = False Then
'        ' la ricetta esiste, ma carica i dati puntati dalla ricetta se
'        ' i puntatori sono validi
'
'        PrevId = Order.Tube.TubeId
'        Order.Tube.TubeId = RecipeData.Recordset.Fields("TubeId")
'        If Not OrdersForm.LoadTubeData(Order.Tube) Then Order.Tube.TubeId = PrevId
'
'        PrevId = Order.Bundle.BundleId
'        Order.Bundle.BundleId = RecipeData.Recordset.Fields("BundleId")
'        If Not OrdersForm.LoadBundleData(Order.Bundle) Then Order.Bundle.BundleId = PrevId
'
'        PrevId = Order.Strap.StrapId
'        Order.Strap.StrapId = RecipeData.Recordset.Fields("StrapId")
'        If Not OrdersForm.LoadStrapData(Order.Strap) Then Order.Strap.StrapId = PrevId
'
'        ' si aggiorna la visualizzazione dei dati appena caricati da ricetta
'        TubeDimensioDisplayUpdate
'        BundleTabUpdate
'        StrapTabUpdate
'
'    Else
'        ' il nome della ricetta non esiste
'    End If

End Sub





            'Private Sub RecipeDeleteCommand_Click()
            '    On Error Resume Next
            '    RecipeData.Recordset.FindFirst ("Item = '" & ItemDisplay.Text & "'")
            '    If RecipeData.Recordset.NoMatch = False Then
            '        If MsgBox(Param.Text("AreYouSure"), vbYesNo, Param.Text("DeleteItem") & " " & ItemDisplay.Text) = vbYes Then
            '            RecipeData.Recordset.Delete
            '        End If
            '    End If
            'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
'    If SSTab1.Tab = 2 And MainMDIForm.StrapEnableFlag = False Then
'        SSTab1.Tab = PreviousTab
'    End If
'
'    If SSTab1.Tab = 1 Then
'        TubeDimensioDisplayUpdate
'    End If
'
'    ' salvataggio ricetta quando si ritorna sul tab 0
'    If SSTab1.Tab = 0 Then
'        CheckData
'        'posiziona puntatore ricetta e aggiorna display ricetta
'        RecipeData.Recordset.FindFirst ("Item = '" & Order.Item & "'")
'        ItemDisplay.Text = Order.Item
'    End If

End Sub


'********************************************************************
' Funzioni di risposta ai comandi dell'operatore in zona tubo e Pacco
'********************************************************************
' forma tubo
Private Sub RoundTubeOption_Click()
    If FromOperator Then
        ' assegna le nuove misure al tubo
        SetTubeShape
        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
        LoadPrevRecipes
        ' aggiorna immagine
        BundleTabUpdate
        ' aggiorna dimensioni tubo
        TubeDimensioDisplayUpdate
    End If
End Sub
Private Sub SquareTubeOption_Click()
    If FromOperator Then
        ' assegna le nuove misure al tubo
        SetTubeShape
        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
        LoadPrevRecipes
        ' aggiorna immagine
        BundleTabUpdate
        ' aggiorna dimensioni tubo
        TubeDimensioDisplayUpdate
    End If
End Sub
Private Sub RectTubeOption_Click()
    If FromOperator Then
        ' assegna le nuove misure al tubo
        SetTubeShape
        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
        LoadPrevRecipes
        ' aggiorna immagine
        BundleTabUpdate
        ' aggiorna dimensioni tubo
        TubeDimensioDisplayUpdate
    End If
End Sub
Private Sub SetTubeShape()
'    SpecialBundleFlag = False
'    Order.Tube.Round = RoundTubeOption.value
'    If (SquareTubeOption.value Or RoundTubeOption.value) Then Order.Tube.Width = Order.Tube.Height
'    ' eco su forma tubo della forma pacco
'    Order.Bundle.RoundTube = Order.Tube.Round
'    If Order.Tube.Round = False Then Order.Bundle.Hex = False
End Sub

' dimensioni tubo
Private Sub LengthDisplay_Change()
'    Dim QueryString As String
'    Dim ret As Boolean
'    If FromOperator Then
'        ' assegna le nuove misure al tubo
'        Order.Tube.Length = Unit.Display_mm_To_m(LengthDisplay.Text, Order.Tube.Length)
'        Order.Strap.TubeLength = Order.Tube.Length
'        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
'        If LoadPrevRecipes = False Then
'            ' se non si è trovata una reggia mediante autoapprendimento
'            ' allora si usa il buon senso e si cerca una reggia con misure
'            ' prossime a quelle del tubo
'            QueryString = "TubeLength <= " & Str(Order.Tube.Length + LengthWindow)
'            StrapData.Recordset.FindLast (QueryString)
'            If StrapData.Recordset.NoMatch Then
'                ' non esisono altre ricette accettabili
'            Else
'                Order.Strap.StrapId = StrapData.Recordset.Fields("StrapId")
'                OrdersForm.LoadStrapData Order.Strap
'            End If
'        End If
'        ' aggiorna immagine
'        BundleTabUpdate
'    End If
End Sub
Private Sub HeightDisplay_Change()
'    If FromOperator Then
'        ' assegna le nuove misure al tubo
'        Order.Tube.Height = Unit.Display_mm_To_m(HeightDisplay.Text, Order.Tube.Height)
'        If SquareTubeOption.value Or RoundTubeOption.value Then Order.Tube.Width = Order.Tube.Height
'        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
'        LoadPrevRecipes
'        ' aggiorna immagine
'        BundleTabUpdate
'    End If
End Sub

Private Sub ThickDisplay_Change()
'    If FromOperator Then
'        ' assegna le nuove misure al tubo
'        Order.Tube.Thickness = Unit.Display_mm_To_m(ThickDisplay.Text, Order.Tube.Thickness)
'        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
'        LoadPrevRecipes
'        ' aggiorna immagine
'        BundleTabUpdate
'    End If
End Sub


Private Sub WidthDisplay_Change()
'    If FromOperator Then
'        ' assegna le nuove misure al tubo
'        Order.Tube.Width = Unit.Display_mm_To_m(WidthDisplay.Text, Order.Tube.Width)
'        ' tenta di assegnare il giusto pacco e reggia al nuovo tubo
'        LoadPrevRecipes
'        ' aggiorna immagine
'        BundleTabUpdate
'    End If
End Sub

' tentativo di caricamento ricette pacco , reggia  e trasporto ad ogni modifica dimensioni tubo
Private Function LoadPrevRecipes() As Boolean
'    Dim PrevId As Long
'    Dim TestOrder As OrderClass
'    Set TestOrder = New OrderClass
'
'    If OrdersForm.GetPrevRecipesID(Order.Tube) Then
'        TestOrder.Bundle.BundleId = Order.Tube.PrevBundleId
'        If OrdersForm.LoadBundleData(TestOrder.Bundle) Then
'            ' esiste una ricetta pacco adatta al tubo , verifica se il tipo di pacco è ammesso
'            If TestOrder.Bundle.Hex And Param.Bit("HexBundle") Then
'                ' pacco ammesso, carica in Order il nuovo pacco
'                Order.Bundle.BundleId = TestOrder.Bundle.BundleId
'                OrdersForm.LoadBundleData Order.Bundle
'            Else
'                If (TestOrder.Bundle.Hex = False) And Param.Bit("SquareBundle") Then
'                    ' pacco ammesso, carica in Order il nuovo pacco
'                    Order.Bundle.BundleId = TestOrder.Bundle.BundleId
'                    OrdersForm.LoadBundleData Order.Bundle
'                Else
'                    ' pacco non ammesso
'                End If
'            End If
'        Else
'            ' ricetta pacco non esistente
'        End If
'
'
'        PrevId = Order.Strap.StrapId
'        Order.Strap.StrapId = Order.Tube.PrevStrapId
'        If OrdersForm.LoadStrapData(Order.Strap) = False Then Order.Strap.StrapId = PrevId
'        LoadPrevRecipes = True
'    Else
'        LoadPrevRecipes = False
'    End If
End Function


' forma pacco
Private Sub SquareBundleOption_Click()
    If FromOperator Then
        SetBundleShape
        BundleTabUpdate
        TubeDimensioDisplayUpdate
    End If
End Sub
Private Sub HexOption_Click()
    If FromOperator Then
        SetBundleShape
        BundleTabUpdate
        TubeDimensioDisplayUpdate
    End If
End Sub
Private Sub SetBundleShape()
'    SpecialBundleFlag = False
'    Order.Bundle.Hex = HexOption.value
'    If HexOption.value Then Order.Bundle.RoundTube = True
'    ' eco su forma tubo della forma pacco
'    Order.Tube.Round = Order.Bundle.RoundTube
'    If Order.Tube.Round Then Order.Tube.Width = Order.Tube.Height
End Sub


' reperisce ricetta con numero tubi superiore
Private Sub IncTubeCommand_Click()
'    Dim ret As Boolean
'    If BundleData.Recordset.EOF = False Then
'        BundleData.Recordset.MoveNext
'        If BundleData.Recordset.EOF = False Then
'            Order.Bundle.BundleId = BundleData.Recordset.Fields("BundleId")
'            ret = OrdersForm.LoadBundleData(Order.Bundle)
'            BundleTabUpdate
'        Else
'            BundleData.Recordset.MovePrevious
'        End If
'    End If
End Sub

' reperisce ricetta con numero tubi inferiore
Private Sub DecTubeCommand_Click()
'    Dim ret As Boolean
'    If BundleData.Recordset.BOF = False Then
'        BundleData.Recordset.MovePrevious
'        If BundleData.Recordset.BOF = False Then
'            Order.Bundle.BundleId = BundleData.Recordset.Fields("BundleId")
'            ret = OrdersForm.LoadBundleData(Order.Bundle)
'            BundleTabUpdate
'        Else
'            BundleData.Recordset.MoveNext
'        End If
'    End If
End Sub

' apre la finestra pacco speciale
Private Sub SpecialBudleCommand_Click()
    SpecialBundleFlag = True
    BundleTabUpdate
End Sub

' se l'operatore tenta di cambiare il numero di tubi o il peso nel pacco si apre la finestra pacco speciale
Private Sub TubesDisplay_Change()
    If FromOperator Then
        SpecialBundleFlag = True
        BundleTabUpdate
    End If
End Sub
Private Sub WeightDisplay_Change()
    If FromOperator Then
        SpecialBundleFlag = True
        BundleTabUpdate
    End If
End Sub

' esce da finestra modifica pacco speciale
Private Sub CancelSpecialCommand_Click()
    If SpecialBundleFlag Then
        SpecialBundleFlag = False
        BundleTabUpdate
    End If
End Sub

' cancellazione pacco speciale dalle ricette
Private Sub DeleteSpecialCommand_Click()
'    '1) cancellazione pacco
'    OrdersForm.DeleteBundleData Order.Bundle
'    '2) posizionamento su pacco precedente nell'archivio
'    BundleData.Recordset.MovePrevious
'    If BundleData.Recordset.BOF Then BundleData.Recordset.MoveNext
'    '3) carica nuovi dati
'    Order.Bundle.BundleId = BundleData.Recordset.Fields("BundleId")
'    '4) aggiorna immagine
'    BundleTabUpdate
End Sub

' aggiunta nuovo pacco speciale alle ricette
Private Sub SaveSpecialBundleCommand_Click()
'    Dim i As Integer
'    '1) si marca il pacco come speciale
'    Order.Bundle.Regular = False
'    '2) si trasferiscono i dati da caselle di testo a dati pacco
'    On Error Resume Next
'        For i = 0 To (MAX_ROWS - 1)
'            If RowEdit(i).Text = "" Then
'                Order.Bundle.TubesRow(i) = 0
'            Else
'                Order.Bundle.TubesRow(i) = CInt(RowEdit(i).Text)
'            End If
'        Next i
'    On Error GoTo 0
'    Order.Bundle.DataCheck  ' controllo composizione pacco
'    '3) si salvano i dati
'    OrdersForm.SaveBundleData Order.Bundle
'    '4) si chiude la finestra pacco speciale
'    SpecialBundleFlag = False
'    '5) si aggiorna l'immagine
'    BundleTabUpdate
End Sub


'********************************************************************
' Funzioni di risposta ai comandi dell'operatore in zona reggiatura
'********************************************************************
Private Sub PosizioneDisplay_Change()
'    If FromOperator Then
'        Order.Strap.Posizione = Unit.Display_mm_To_m(PosizioneDisplay.Text, Order.Strap.Posizione)
'        If Order.Strap.Posizione > 10# Then
'            Order.Strap.Posizione = 10#
'        End If
'        If Order.Strap.Posizione < 1# Then
'            Order.Strap.Posizione = 1#
'        End If
'    End If
End Sub

Private Sub StrapsDisplay_Change()
'    If FromOperator Then
'        SpecialStrapFlag = True
'        Order.Strap.AutomaticQuotes (Unit.Display_int_To_int(StrapsDisplay.Text, Order.Strap.StrapNumber))
'        StrapConfigShow
'    End If
End Sub

Private Sub FirstDisplay_Change()
'    If FromOperator Then
'        SpecialStrapFlag = True
'        Order.Strap.Quote(0) = Unit.Display_mm_To_m(FirstDisplay.Text, Order.Strap.Quote(0))
'        Order.Strap.AutomaticQuotes (Order.Strap.StrapNumber)
'        StrapConfigShow
'    End If
End Sub

Private Sub IncStrapCommand_Click()
'    Dim ret As Boolean
'    Dim QueryString As String
'    If StrapData.Recordset.EOF = False Then
'        QueryString = "TubeLength > " & Str(Order.Tube.Length - LengthWindow) & " AND TubeLength < " & Str(Order.Tube.Length + LengthWindow)
'        StrapData.Recordset.FindNext (QueryString)
'        If StrapData.Recordset.NoMatch Then
'            ' non esisono altre ricette
'        Else
'            Order.Strap.StrapId = StrapData.Recordset.Fields("StrapId")
'            ret = OrdersForm.LoadStrapData(Order.Strap)
'        End If
'        StrapTabUpdate
'    End If
End Sub
Private Sub DecStrapCommand_Click()
'    Dim ret As Boolean
'    Dim QueryString As String
'    If StrapData.Recordset.BOF = False Then
'        QueryString = "TubeLength > " & Str(Order.Tube.Length - LengthWindow) & " AND TubeLength < " & Str(Order.Tube.Length + LengthWindow)
'        StrapData.Recordset.FindPrevious (QueryString)
'        If StrapData.Recordset.NoMatch Then
'            ' non esisono altre ricette
'        Else
'            Order.Strap.StrapId = StrapData.Recordset.Fields("StrapId")
'            ret = OrdersForm.LoadStrapData(Order.Strap)
'        End If
'        StrapTabUpdate
'    End If
End Sub

Private Sub SpecialStrapCommand_Click()
    SpecialStrapFlag = True
    StrapTabUpdate
End Sub
Private Sub SaveSpecialStrapCommand_Click()
'    Dim i As Integer
'    '1) si marca la reggia come speciale
'    Order.Strap.Regular = False
'    '2) si trasferiscono i dati da caselle di testo a dati reggie
'    Order.Strap.TubeLength = Order.Tube.Length
'    For i = 0 To (StrapsDisplay.Text - 1)
'        Order.Strap.Quote(i) = Unit.Display_mm_To_m(QuoteEdit(i).Text, Order.Strap.Quote(i))
'    Next i
'    For i = StrapsDisplay.Text To (MAX_STRAPS - 1)
'        Order.Strap.Quote(i) = 0#
'    Next i
'    '3) si salvano i dati
'    OrdersForm.SaveStrapData Order.Strap
'    '4) si chiude la finestra pacco speciale
'    SpecialStrapFlag = False
'    '5) si aggiorna l'immagine
'    StrapTabUpdate
End Sub

Private Sub DeleteSpecialStrapCommand_Click()
'    '1) cancellazione pacco
'    OrdersForm.DeleteStrapData Order.Strap
'    '2) posizionamento su pacco precedente nell'archivio
'    StrapData.Recordset.MovePrevious
'    If StrapData.Recordset.BOF Then StrapData.Recordset.MoveNext
'    '3) carica nuovi dati
'    Order.Strap.StrapId = StrapData.Recordset.Fields("StrapId")
'    '4) aggiorna immagine
'    StrapTabUpdate
End Sub

Private Sub CancelSpecialStrapCommand_Click()
    If SpecialStrapFlag Then
        SpecialStrapFlag = False
        StrapTabUpdate
    End If
End Sub



'***********************************************************
'        Recupera i dati in zona etichetta e cambio commessa
'***********************************************************
Private Sub LabelTabUpload()
'    ' trasferimento dati cartellino
'    Dim TotString, OneString As String
'    Dim i As Integer
'    For i = 0 To (PrintFieldNumber - 1)
'        On Error Resume Next ' per saltare senza errori i campi ad agg. autom non esistenti in questa finestra
'        ' carica su OneString il campo editato
'        ' troncandolo se troppo lungo o
'        ' aggiungendo spazi finali se troppo corto
'        OneString = StringFieldLength
'        LSet OneString = PrintEdit(i)
'        ' accoda il campo così formattato nella stringa dei
'        ' dati del cartellino
'        TotString = TotString & OneString
'    Next i
'    Order.PrinterData = TotString
'
'    ' trasferimento dati display ordine :
'    ' primo campo (Trim=elimina spazi iniziali e finali)
'    TotString = Trim(PrintEdit(Param.Number("PosizCampoOrdineVisibile_1")))
'    If Param.Number("NumeroCampiVisibiliOrdine") > 1 Then
'        ' secondo campo
'        OneString = Trim(PrintEdit(Param.Number("PosizCampoOrdineVisibile_2")))
'        TotString = TotString & DisplaySeparator & OneString
'    End If
'    If Param.Number("NumeroCampiVisibiliOrdine") > 2 Then
'        ' terzo campo
'        OneString = Trim(PrintEdit(Param.Number("PosizCampoOrdineVisibile_3")))
'        TotString = TotString & DisplaySeparator & OneString
'    End If
'    Order.DisplayData = TotString
'
'    ' dati quantità programmata
'    If ContOption.value Then
'        Order.AutomaticChange = False
'        Order.AutomaticStop = False
'    Else
'        Order.AutomaticChange = AutOption.value
'        Order.AutomaticStop = StopOption.value
'    End If
'
'    ' nome ricetta
'    Order.Item = ItemDisplay.Text
'
'    On Error Resume Next
'    Order.BundlesPreset = CInt(OrderBundles.Text)
End Sub


'*****************************************************************************
'    funzioni ausiliarie per aggiornamento display etichetta e cambio commessa
'*****************************************************************************
' Aggiorna tab dati quantità ed etichetta
Private Sub LabelTabUpdate()
'    ' disabilitazione richiamo PictureUpdate dall'interno degli eventi Click_
'    ' generati dalla PictureUpdate stessa
'    FromOperator = False
'
'    On Error Resume Next
'    ' dati etichetta
'    Dim i As Integer
'    For i = 0 To (PrintFieldNumber - 1)
'        PrintEdit(i).Text = RTrim(Mid(Order.PrinterData, 1 + PrintFieldLength * i, PrintFieldLength))
'    Next i
'    ' quantità
'    OrderBundles.Text = Order.BundlesPreset
'    ' tipo cambio ordine
'    ContOption.value = (Not Order.AutomaticChange) And (Not Order.AutomaticStop)
'    AutOption.value = Order.AutomaticChange
'    StopOption.value = Not Order.AutomaticChange
'
'    'posiziona puntatore ricetta e aggiorna display ricetta
'    RecipeData.Refresh
'    RecipeData.Recordset.FindFirst ("Item = '" & Order.Item & "'")
'    ItemDisplay.Text = Order.Item
'
'    '******* Ripristino flag di modifica dati *****************
'    FromOperator = True
'    '**********************************************************

End Sub


'*************************************************************
'    Inizio funzioni ausiliarie per disegno pacco
'*************************************************************
' ridisegno tab pacco con esclusione delle misure del tubo
Private Sub BundleTabUpdate()
'    Dim i As Integer
'    ' disabilitazione richiamo PictureUpdate dall'interno degli eventi Click_
'    ' generati dalla PictureUpdate stessa
'    FromOperator = False
'
'
'    ' inizializza tabella e record corrente dell'oggetto data
'    If Order.Bundle.Hex Then
'        BundleData.RecordSource = "HexBundle"
'    Else
'        If Order.Bundle.RoundTube Then
'            BundleData.RecordSource = "SqRdBundle"
'        Else
'            BundleData.RecordSource = "SqSqBundle"
'        End If
'    End If
'    BundleData.Refresh
'    BundleData.Recordset.FindFirst ("BundleId=" & CStr(Order.Bundle.BundleId))
'    If BundleData.Recordset.NoMatch Then BundleData.Recordset.MoveFirst
'    Order.Bundle.BundleId = BundleData.Recordset.Fields("BundleId")
'    OrdersForm.LoadBundleData Order.Bundle
'
'    ' Display forma tubo
'    TubeShapeDraw
'
'    ' forma pacco
'    If Order.Bundle.Hex Then
'        HexOption.value = True
'    Else
'        SquareBundleOption.value = True
'    End If
'    ' totale tubi
'    TubesDisplay.Text = Order.Bundle.Tubes
'    ' peso pacco teorico
'    WeightDisplay.Text = Unit.kg_To_Display_kg(Order.Tube.Weight * Order.Bundle.Tubes)
'
'    If SpecialBundleFlag Then
'        ' visualizza dati pacco speciale
'        PackDrawFrame.Visible = False
'        SpecialFrame.Visible = True
'        ' dati pacco speciale
'        For i = 0 To (MAX_ROWS - 1)
'            If Order.Bundle.TubesRow(i) > 0 Then
'                RowEdit(i).Text = Order.Bundle.TubesRow(i)
'            Else
'                RowEdit(i).Text = ""
'            End If
'        Next i
'    Else
'        ' visualizza disegno pacco
'        BundleDraw
'        PackDrawFrame.Visible = True
'        SpecialFrame.Visible = False
'        If Order.Bundle.Regular Then
'            DeleteSpecialCommand.Visible = False
'        Else
'            DeleteSpecialCommand.Visible = True
'        End If
'    End If
'
'
'    '******* Ripristino flag di modifica dati *****************
'    FromOperator = True
'    '**********************************************************

End Sub

'ridisegno misure tubo
Private Sub TubeDimensioDisplayUpdate()
'    ' disabilitazione richiamo PictureUpdate dall'interno degli eventi Click_
'    ' generati dalla PictureUpdate stessa
'    FromOperator = False
'
'    LengthDisplay.Text = Unit.m_To_Display_mm(Order.Tube.Length)
'    HeightDisplay.Text = Unit.m_To_Display_mm0(Order.Tube.Height)
'    WidthDisplay.Text = Unit.m_To_Display_mm0(Order.Tube.Width)
'    ThickDisplay.Text = Unit.m_To_Display_mm00(Order.Tube.Thickness)
'
'    '******* Ripristino flag di modifica dati *****************
'    FromOperator = True
'    '**********************************************************
End Sub

' ridisegno completo del pacco ad eccezione dei display di edit
Private Sub BundleDraw()
    ' disegno pacco
    UpdateTubesPosition
    UpdateDrawData
    QuoteDraw
    TubesDraw
End Sub

' aggiornamento configurazione display misure in funzione della forma del tubo
Private Sub TubeShapeDraw()
'    If Order.Bundle.RoundTube Then
'        RoundTubeOption.value = True
'        DiameterLabel.Visible = True
'        DimensionLabel.Visible = False
'        XLabel.Visible = False
'        HeightDisplay.Visible = True
'        WidthDisplay.Visible = False
'    Else
'        DiameterLabel.Visible = False
'        DimensionLabel.Visible = True
'        XLabel.Visible = True
'        HeightDisplay.Visible = True
'        WidthDisplay.Visible = True
'        If RectTubeOption.value Then
'            WidthDisplay.Enabled = True
'        Else
'            If Order.Tube.Width = Order.Tube.Height Then
'                SquareTubeOption.value = True
'                WidthDisplay.Enabled = False
'            Else
'                RectTubeOption.value = True
'                WidthDisplay.Enabled = True
'            End If
'        End If
'    End If
End Sub


' posizionamento tubi in coordinate reali (m)
' con centro pacco in posizione ( 0 , 0 )
Private Sub UpdateTubesPosition()
    Dim BundleTubeCounter, RowTubeCounter, RowCounter As Integer
'    Dim Rows As Integer
'    Dim Height As Double
'    Dim Width As Double
'
'    BundleTubeCounter = 0
'    RowCounter = 0
'    With Order
'        ' carica variabili locali per evitare continui richiami di funzione
'        Rows = .Bundle.Rows
'        Height = .Bundle.Height(.Tube.Height)
'        Width = .Bundle.Width(.Tube.Width)
'        While (RowCounter < Rows) And (RowCounter < MAX_ROWS)
'            ' posizionamento coordinate X e Y del primo tubo di ogni fila
'            If .Bundle.RoundTube Then
'                RTubeX(BundleTubeCounter) = .Tube.Width * ((.Bundle.TubesRow(RowCounter) - 1#) / -2#)
'                If RowCounter = 0 Then
'                    ' primo tubo, prima fila
'                    RTubeY(BundleTubeCounter) = (Height - .Tube.Height) * 0.5
'                Else
'                    ' primo tubo delle altre file
'                    If (.Bundle.TubesRow(RowCounter) <> .Bundle.TubesRow(RowCounter - 1)) Then
'                        RTubeY(BundleTubeCounter) = RTubeY(BundleTubeCounter - 1) - .Tube.Height * 0.866025
'                    Else
'                        RTubeY(BundleTubeCounter) = RTubeY(BundleTubeCounter - 1) - .Tube.Height
'                    End If
'                End If
'            Else
'                RTubeX(BundleTubeCounter) = .Tube.Width * ((.Bundle.TubesRow(RowCounter) - 1#) / -2#)
'                If Rows > 1 Then
'                    RTubeY(BundleTubeCounter) = (Height - .Tube.Height) * 0.5 - RowCounter * (Height / Rows)
'                Else
'                    RTubeY(BundleTubeCounter) = (Height - .Tube.Height) * 0.5
'                End If
'            End If
'            BundleTubeCounter = BundleTubeCounter + 1
'            If BundleTubeCounter >= MAX_TUBES Then BundleTubeCounter = MAX_TUBES
'
'            ' posizionamento tubi successivi nella fila
'            RowTubeCounter = 1
'            While RowTubeCounter < .Bundle.TubesRow(RowCounter)
'                RTubeX(BundleTubeCounter) = RTubeX(BundleTubeCounter - 1) + .Tube.Width
'                RTubeY(BundleTubeCounter) = RTubeY(BundleTubeCounter - 1)
'                BundleTubeCounter = BundleTubeCounter + 1
'                If BundleTubeCounter >= MAX_TUBES Then BundleTubeCounter = MAX_TUBES
'                RowTubeCounter = RowTubeCounter + 1
'            Wend
'            RowCounter = RowCounter + 1
'        Wend
'        ' prepara quotature
'        RQuoteX(0) = .Bundle.Base(.Tube.Width) * -0.5: RQuoteY(0) = Height * 0.5
'        RQuoteX(1) = .Bundle.Base(.Tube.Width) * 0.5: RQuoteY(1) = Height * 0.75
'        RQuoteX(2) = Width * 0.55: RQuoteY(2) = Height * -0.5
'        RQuoteX(3) = Width * -0.5: RQuoteY(3) = 0#
'        RQuoteX(4) = RQuoteX(3) - Width * 0.17: RQuoteY(4) = RQuoteY(3) + Width * 0.1
'        RQuoteX(5) = RQuoteX(0) - Width * 0.17: RQuoteY(5) = RQuoteY(0) + Width * 0.1
'    End With
End Sub

' calcola posizione tubi e quote in pixel
Private Sub UpdateDrawData()
'    ' calcolo fattore di scala
'    Dim Coeff, XCoeff, YCoeff As Double
'    Dim Xoffset, Yoffset As Integer
'    Dim i As Integer
'
'    If Order.Bundle.Hex Then
'        If RQuoteX(4) <> 0 Then XCoeff = ((PackDrawFrame.Width - HeightLabel.Width)) / (RQuoteX(4) * -2#)
'    Else
'        If RQuoteX(2) <> 0 Then XCoeff = ((PackDrawFrame.Width - HeightLabel.Width)) / (RQuoteX(2) * 2#)
'    End If
'
'    If RQuoteY(1) <> 0 Then YCoeff = PackDrawFrame.Height / (RQuoteY(1) * 2#)
'    If XCoeff > YCoeff Then
'        Coeff = YCoeff
'    Else
'        Coeff = XCoeff
'    End If
'    Coeff = Coeff * 0.95
'
'    ' calcolo spostamento per centratura
'    If Order.Bundle.Hex Then
'        Xoffset = PackDrawFrame.Width / 2
'    Else
'        Xoffset = (PackDrawFrame.Width - HeightLabel.Width) / 2
'    End If
'    Yoffset = PackDrawFrame.Height * 0.45
'
'    ' scalamento e centratura
'    PTubeHeight = Order.Tube.Height * Coeff 'altezza o diametro tubo
'    PTubeWidth = Order.Tube.Width * Coeff   'larghezza tubo o lato lungo profilo
'
'    For i = 0 To 5
'        PQuoteX(i) = RQuoteX(i) * Coeff + Xoffset
'        PQuoteY(i) = RQuoteY(i) * Coeff + Yoffset
'    Next i
'    For i = 0 To MAX_TUBES
'        PTubeX(i) = RTubeX(i) * Coeff + Xoffset
'        PTubeY(i) = RTubeY(i) * Coeff + Yoffset
'    Next i
End Sub

Private Sub MainQuoteDraw()
'    ' **************** Quotatura base pacco ***************
'    ' prima linea
'    With ArrowLine(0): .X1 = PQuoteX(0): .Y1 = PQuoteY(0): .X2 = PQuoteX(0): .Y2 = PQuoteY(1):    End With
'    'seconda linea
'    With ArrowLine(1): .X1 = PQuoteX(0): .Y1 = PQuoteY(1): .X2 = PQuoteX(1): .Y2 = PQuoteY(1):    End With
'    'terza linea
'    With ArrowLine(2): .X1 = PQuoteX(1): .Y1 = PQuoteY(1): .X2 = PQuoteX(1): .Y2 = PQuoteY(0): End With
'    ' prima freccia
'    With ArrowLine(3): .X1 = PQuoteX(0): .Y1 = PQuoteY(1): .X2 = PQuoteX(0) + ArrowLength: .Y2 = PQuoteY(1) - ArrowWidth:    End With
'    With ArrowLine(4): .X1 = PQuoteX(0): .Y1 = PQuoteY(1): .X2 = PQuoteX(0) + ArrowLength: .Y2 = PQuoteY(1) + ArrowWidth:    End With
'    'seconda freccia
'    With ArrowLine(5): .X1 = PQuoteX(1): .Y1 = PQuoteY(1): .X2 = PQuoteX(1) - ArrowLength: .Y2 = PQuoteY(1) - ArrowWidth:    End With
'    With ArrowLine(6): .X1 = PQuoteX(1): .Y1 = PQuoteY(1): .X2 = PQuoteX(1) - ArrowLength: .Y2 = PQuoteY(1) + ArrowWidth:    End With
'    'testo
'    BaseLabel.Caption = Unit.m_To_Display_mm(Order.Bundle.Base(Order.Tube.Width)) & Unit.mmString
'    BaseLabel.Left = ((PQuoteX(0) + PQuoteX(1)) / 2) - (BaseLabel.Width / 2)
'    BaseLabel.Top = PQuoteY(1) - BaseLabel.Height
'
'    ' **************** Quotatura altezza pacco ************
'    ' prima linea
'    With ArrowLine(7): .X1 = PQuoteX(1): .Y1 = PQuoteY(0): .X2 = PQuoteX(2): .Y2 = PQuoteY(0):    End With
'    'seconda linea
'    With ArrowLine(8): .X1 = PQuoteX(2): .Y1 = PQuoteY(2): .X2 = PQuoteX(1): .Y2 = PQuoteY(2):    End With
'    'terza linea
'    With ArrowLine(9): .X1 = PQuoteX(2): .Y1 = PQuoteY(0): .X2 = PQuoteX(2): .Y2 = PQuoteY(2): End With
'    ' prima freccia
'    With ArrowLine(10): .X1 = PQuoteX(2): .Y1 = PQuoteY(0): .X2 = PQuoteX(2) - ArrowWidth: .Y2 = PQuoteY(0) - ArrowLength:    End With
'    With ArrowLine(11): .X1 = PQuoteX(2): .Y1 = PQuoteY(0): .X2 = PQuoteX(2) + ArrowWidth: .Y2 = PQuoteY(0) - ArrowLength:    End With
'    'seconda freccia
'    With ArrowLine(12): .X1 = PQuoteX(2): .Y1 = PQuoteY(2): .X2 = PQuoteX(2) - ArrowWidth: .Y2 = PQuoteY(2) + ArrowLength:    End With
'    With ArrowLine(13): .X1 = PQuoteX(2): .Y1 = PQuoteY(2): .X2 = PQuoteX(2) + ArrowWidth: .Y2 = PQuoteY(2) + ArrowLength:    End With
'    'testo
'    HeightLabel.Caption = Unit.m_To_Display_mm(Order.Bundle.Height(Order.Tube.Height)) & Unit.mmString
'    HeightLabel.Left = PQuoteX(2)
'    HeightLabel.Top = ((PQuoteY(2) + PQuoteY(0)) / 2) - (HeightLabel.Height / 2)
'
'    ' nasconde gli oggetti di quotatura non utilizzati
'    SideLabel.Visible = False
'    ArrowLine(14).Visible = False
'    ArrowLine(15).Visible = False
'    ArrowLine(16).Visible = False
'    ArrowLine(17).Visible = False
'    ArrowLine(18).Visible = False
'    ArrowLine(19).Visible = False
'    ArrowLine(20).Visible = False
End Sub

Private Sub SideQuoteDraw()
'    SideLabel.Visible = True
'    ArrowLine(14).Visible = True
'    ArrowLine(15).Visible = True
'    ArrowLine(16).Visible = True
'    ArrowLine(17).Visible = True
'    ArrowLine(18).Visible = True
'    ArrowLine(19).Visible = True
'    ArrowLine(20).Visible = True
'    ' prima linea
'    With ArrowLine(14): .X1 = PQuoteX(3): .Y1 = PQuoteY(3): .X2 = PQuoteX(4): .Y2 = PQuoteY(4):    End With
'    'seconda linea
'    With ArrowLine(15): .X1 = PQuoteX(4): .Y1 = PQuoteY(4): .X2 = PQuoteX(5): .Y2 = PQuoteY(5):    End With
'    'terza linea
'    With ArrowLine(16): .X1 = PQuoteX(5): .Y1 = PQuoteY(5): .X2 = PQuoteX(0): .Y2 = PQuoteY(0):    End With
'    ' prima freccia
'    With ArrowLine(17): .X1 = PQuoteX(4): .Y1 = PQuoteY(4): .X2 = PQuoteX(4): .Y2 = PQuoteY(4) + ArrowLength:   End With
'    With ArrowLine(18): .X1 = PQuoteX(4): .Y1 = PQuoteY(4): .X2 = PQuoteX(4) + ArrowLength: .Y2 = PQuoteY(4) + ArrowLength * 0.5:     End With
'    'seconda freccia
'    With ArrowLine(19): .X1 = PQuoteX(5): .Y1 = PQuoteY(5): .X2 = PQuoteX(5): .Y2 = PQuoteY(5) - ArrowLength:    End With
'    With ArrowLine(20): .X1 = PQuoteX(5): .Y1 = PQuoteY(5): .X2 = PQuoteX(5) - ArrowLength: .Y2 = PQuoteY(5) - ArrowLength * 0.5:    End With
'    'testo
'    SideLabel.Caption = Unit.m_To_Display_mm(Order.Bundle.Side(Order.Tube.Height)) & Unit.mmString
'    SideLabel.Left = (PQuoteX(5) + PQuoteX(4) - SideLabel.Width) / 2
'    SideLabel.Top = (PQuoteY(5) + PQuoteY(4) - SideLabel.Height) / 2
'    SideLabel.Visible = True
End Sub

' disegno quote
Private Sub QuoteDraw()
'    MainQuoteDraw
'    If Order.Bundle.Hex Then SideQuoteDraw
End Sub

' disegno tubi pacco
Private Sub TubesDraw()
'    ' limitazione numero tubi
'    Dim TotalTubes As Integer
'    TotalTubes = Order.Bundle.Tubes
'    ' disegno tubi
'    Dim i As Integer
'    Dim ShapeHeight, ShapeWidth, ShapeLeft, ShapeTop As Integer
'
'    ShapeHeight = PTubeHeight - LineWidthTwip
'    ShapeWidth = PTubeWidth - LineWidthTwip
'    ' il minimo valore di dimensione in twips è 15
'    If ShapeHeight < 15 Then ShapeHeight = 15
'    If ShapeWidth < 15 Then ShapeWidth = 15
'
'    ShapeLeft = ShapeWidth / 2
'    ShapeTop = ShapeHeight / 2
'
'    For i = 0 To MAX_TUBES
'        TubeShape(i).Visible = False
'        If i < TotalTubes Then
'            TubeShape(i).BorderWidth = LineWidthPixel ' spessore linea
'            If Order.Bundle.RoundTube Then
'                TubeShape(i).Shape = 3 ' round
'            Else
'                TubeShape(i).Shape = 4 ' rectangle
'            End If
'            'TubeShape(i).BorderColor = VirtualTubeColor
'            TubeShape(i).BorderColor = PresentTubeColor
'            TubeShape(i).Height = ShapeHeight
'            TubeShape(i).Width = ShapeWidth
'            TubeShape(i).Left = PTubeX(i) - ShapeLeft
'            TubeShape(i).Top = PTubeY(i) - ShapeTop
'            TubeShape(i).Visible = True
'        End If
'    Next i
End Sub

'*************************************************************
'    Fine funzioni ausiliarie per disegno pacco
'*************************************************************




'*************************************************************
'    Inizio funzioni ausiliarie per disegno reggie
'*************************************************************
Private Sub StrapTabUpdate()
'    Dim i As Integer
'    ' disabilitazione richiamo PictureUpdate dall'interno degli eventi Click_
'    ' generati dalla PictureUpdate stessa
'    FromOperator = False
'
'
'    ' inizializza tabella e record corrente dell'oggetto data
'    StrapData.Refresh
'    StrapData.Recordset.FindFirst ("StrapId=" & CStr(Order.Strap.StrapId))
'    If StrapData.Recordset.NoMatch Then StrapData.Recordset.MoveFirst
'    Order.Strap.StrapId = StrapData.Recordset.Fields("StrapId")
'    OrdersForm.LoadStrapData Order.Strap
'
'    ' visualizza dati quote reggie
'    FirstDisplay.Text = Unit.m_To_Display_mm(Order.Strap.Quote(0))
'    StrapsDisplay.Text = Order.Strap.StrapNumber
'
'    ' Aggiunta 09.08.2000
'    PosizioneDisplay.Text = Unit.m_To_Display_mm(Order.Strap.Posizione)
'    ' Fine aggiunta 09.08.2000
'
'    StrapConfigShow
'
'    '******* Ripristino flag di modifica dati *****************
'    FromOperator = True
'    '**********************************************************

End Sub

Private Sub StrapConfigShow()
'    If SpecialStrapFlag Then
'        StrapDrawFrame.Visible = False
'        StrapQuotesUpdate
'    Else
'        ' visualizza disegno reggie
'        StrapDrawUpdate
'        StrapDrawFrame.Visible = True
'        If Order.Strap.Regular Then
'            DeleteSpecialStrapCommand.Visible = False
'        Else
'            DeleteSpecialStrapCommand.Visible = True
'        End If
'    End If
End Sub

'aggiornamento display quote di reggiatura
Private Sub StrapQuotesUpdate()
'    Dim i As Integer
'    For i = 0 To (MAX_STRAPS - 1)
'        If Order.Strap.Quote(i) > 0 Then
'            QuoteEdit(i).Text = Unit.m_To_Display_mm(Order.Strap.Quote(i))
'        Else
'            QuoteEdit(i).Text = ""
'        End If
'    Next i
End Sub


Private Sub StrapDrawUpdate()
'    Dim i As Integer
'    TwipOffset = BundleShape.Left '+ BundleShape.Width
'    If Order.Tube.Length > 0 Then
'        TwipCoeff = BundleShape.Width / Order.Tube.Length
'    Else
'        TwipCoeff = 0
'    End If
'    For i = 0 To (MAX_STRAPS - 1)
'        StrapLine(i).Visible = False
'        StrapLine(i).X1 = TwipOffset + Order.Strap.Quote(i) * TwipCoeff
'        StrapLine(i).Y1 = BundleShape.Top
'        StrapLine(i).X2 = StrapLine(i).X1
'        StrapLine(i).Y2 = StrapLine(i).Y1 + BundleShape.Height
'        If Order.Strap.Quote(i) > 0 Then StrapLine(i).Visible = True
'    Next i
End Sub

'*************************************************************
'    Fine funzioni ausiliarie per disegno reggie
'*************************************************************

