VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form OrderModifyForm 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   270
      TabIndex        =   0
      Top             =   10245
      Width           =   2295
   End
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
      Left            =   2625
      TabIndex        =   1
      Top             =   10230
      Width           =   2295
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   285
      TabIndex        =   2
      Top             =   165
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   17595
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "ORDER"
      TabPicture(0)   =   "OrderModifyForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameRecipe"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameOrderEnd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameOrderTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameSetup"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Ticket"
      TabPicture(1)   =   "OrderModifyForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LabelFrame"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame LabelFrame 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Bundle label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8790
         Left            =   -71910
         TabIndex        =   17
         Top             =   720
         Width           =   7755
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   180
            MaxLength       =   12
            TabIndex        =   50
            Text            =   "3"
            Top             =   1620
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   180
            MaxLength       =   12
            TabIndex        =   49
            Text            =   "5"
            Top             =   1965
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   48
            Text            =   "4"
            Top             =   1620
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   47
            Text            =   "6"
            Top             =   1962
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   180
            MaxLength       =   12
            TabIndex        =   46
            Text            =   "7"
            Top             =   2310
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   180
            MaxLength       =   12
            TabIndex        =   45
            Text            =   "9"
            Top             =   2640
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   10
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   44
            Text            =   "10"
            Top             =   2646
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   11
            Left            =   180
            MaxLength       =   12
            TabIndex        =   43
            Text            =   "11"
            Top             =   2985
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   12
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   42
            Text            =   "12"
            Top             =   2988
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   13
            Left            =   180
            MaxLength       =   12
            TabIndex        =   41
            Text            =   "13"
            Top             =   3330
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   14
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   40
            Text            =   "14"
            Top             =   3330
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   15
            Left            =   180
            MaxLength       =   12
            TabIndex        =   39
            Text            =   "15"
            Top             =   3675
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   16
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   38
            Text            =   "16"
            Top             =   3672
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   17
            Left            =   180
            MaxLength       =   12
            TabIndex        =   37
            Text            =   "17"
            Top             =   4020
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   18
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   36
            Text            =   "18"
            Top             =   4014
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   19
            Left            =   180
            MaxLength       =   12
            TabIndex        =   35
            Text            =   "19"
            Top             =   4350
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   20
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   34
            Text            =   "20"
            Top             =   4356
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   21
            Left            =   180
            MaxLength       =   12
            TabIndex        =   33
            Text            =   "21"
            Top             =   4695
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   22
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   32
            Text            =   "22"
            Top             =   4698
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   23
            Left            =   180
            MaxLength       =   12
            TabIndex        =   31
            Text            =   "23"
            Top             =   5040
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   24
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   30
            Text            =   "24"
            Top             =   5040
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   25
            Left            =   180
            MaxLength       =   12
            TabIndex        =   29
            Text            =   "25"
            Top             =   5385
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   26
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   28
            Text            =   "26"
            Top             =   5382
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   27
            Left            =   180
            MaxLength       =   12
            TabIndex        =   27
            Text            =   "27"
            Top             =   5730
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   28
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   26
            Text            =   "28"
            Top             =   5724
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   29
            Left            =   180
            MaxLength       =   12
            TabIndex        =   25
            Text            =   "29"
            Top             =   6060
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   30
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   24
            Text            =   "30"
            Top             =   6066
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   31
            Left            =   180
            MaxLength       =   12
            TabIndex        =   23
            Text            =   "31"
            Top             =   6420
            Width           =   1755
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   32
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   22
            Text            =   "32"
            Top             =   6420
            Width           =   3375
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   180
            MaxLength       =   36
            TabIndex        =   21
            Text            =   "2"
            Top             =   1200
            Width           =   5160
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   180
            MaxLength       =   36
            TabIndex        =   20
            Text            =   "1"
            Top             =   780
            Width           =   5160
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Index           =   0
            Left            =   150
            MaxLength       =   20
            TabIndex        =   19
            Text            =   "0"
            Top             =   105
            Width           =   5160
         End
         Begin VB.TextBox PrintEdit 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   1980
            MaxLength       =   23
            TabIndex        =   18
            Text            =   "8"
            Top             =   2304
            Width           =   3375
         End
      End
      Begin VB.Frame FrameSetup 
         BackColor       =   &H0080FFFF&
         Caption         =   "SETUP"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   7995
         Left            =   4305
         TabIndex        =   16
         Top             =   1830
         Width           =   10440
         Begin VB.TextBox DisplayLunghezza 
            Alignment       =   2  'Center
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
            Height          =   540
            Left            =   8325
            MaxLength       =   6
            TabIndex        =   53
            Text            =   "0"
            Top             =   4725
            Width           =   1650
         End
         Begin dp6.Pacco6 TuboModOrdine 
            Height          =   2895
            Left            =   6585
            TabIndex        =   52
            Top             =   1290
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   5106
         End
         Begin dp6.Pacco6 PaccoModOrdine 
            Height          =   4560
            Left            =   570
            TabIndex        =   51
            Top             =   990
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   8043
         End
         Begin VB.Label LabelLunghezza 
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
            Left            =   5370
            TabIndex        =   54
            Top             =   4665
            Width           =   2340
         End
      End
      Begin VB.Frame FrameOrderTitle 
         BackColor       =   &H000080FF&
         Height          =   1110
         Left            =   210
         TabIndex        =   14
         Top             =   630
         Width           =   14520
         Begin VB.TextBox TitleDisplay 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   3135
            TabIndex        =   15
            Text            =   "100"
            Top             =   180
            Width           =   8490
         End
      End
      Begin VB.Frame FrameOrderEnd 
         BackColor       =   &H0080FFFF&
         Caption         =   "AT ORDER END "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   7980
         Left            =   240
         TabIndex        =   3
         Top             =   1830
         Width           =   3960
         Begin VB.CommandButton FineOrdine_NuovoOrdine 
            Caption         =   "Nuovo ordine"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   502
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   3855
            Width           =   2970
         End
         Begin VB.CommandButton FineOrdine_Arresto 
            Caption         =   "Arresto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   502
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2520
            Width           =   2970
         End
         Begin VB.CommandButton FineOrdine_NonStop 
            Caption         =   "Non stop"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   510
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   645
            Width           =   2970
         End
         Begin VB.TextBox OrderBundles 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   1200
            TabIndex        =   5
            Text            =   "100"
            Top             =   5910
            Width           =   1575
         End
         Begin VB.Label BundlesLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pacchi"
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
            Left            =   1110
            TabIndex        =   7
            Top             =   7050
            Width           =   1755
         End
         Begin VB.Label AfterLabel 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dopo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1260
            TabIndex        =   6
            Top             =   5145
            Width           =   1455
         End
      End
      Begin VB.Frame FrameRecipe 
         BackColor       =   &H0080FFFF&
         Caption         =   "RECIPE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   7935
         Left            =   4350
         TabIndex        =   4
         Top             =   1875
         Width           =   10380
         Begin VB.CommandButton CommandPgUp 
            BackColor       =   &H0000FF00&
            Caption         =   "PgUp"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2000
            Left            =   9105
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1395
            Width           =   1300
         End
         Begin VB.CommandButton CommandPgDn 
            BackColor       =   &H0000FF00&
            Caption         =   "PgDn"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2000
            Left            =   9105
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   5880
            Width           =   1300
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFF00&
            Caption         =   "current "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   120
            TabIndex        =   8
            Top             =   345
            Width           =   10275
            Begin VB.TextBox ItemDisplay 
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
               Height          =   690
               HideSelection   =   0   'False
               Left            =   90
               TabIndex        =   10
               Text            =   "100"
               Top             =   300
               Width           =   8490
            End
            Begin VB.CommandButton RecipeDeleteCommand 
               BackColor       =   &H0000FF00&
               Caption         =   "DELETE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   8715
               MaskColor       =   &H0000FF00&
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   195
               Width           =   1440
            End
         End
         Begin MSDBGrid.DBGrid RecipeDBGrid 
            Bindings        =   "OrderModifyForm.frx":0038
            Height          =   6480
            Left            =   150
            OleObjectBlob   =   "OrderModifyForm.frx":0051
            TabIndex        =   13
            Top             =   1395
            Width           =   8865
         End
      End
   End
End
Attribute VB_Name = "OrderModifyForm"
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

'===============================================================================
'--------------------- INIZIO FUNZIONI PAGINA
'===============================================================================
Private Sub Form_Load()
    Ok = False
    SSTab1.TabCaption(0) = Param.Text("Order")
    SSTab1.TabCaption(1) = Param.Text("Bundle")
    
    If Param.Bit("recipeEnable") Then
        FrameRecipe.Visible = True
    Else
        FrameRecipe.Visible = False
    End If
    



'    FrameOrderEnd.Caption = Param.Text("AtTheEnd")
'    AutOption.Caption = Param.Text("AutomChange")
'    StopOption.Caption = Param.Text("Stop")
'    ContOption.Caption = Param.Text("ContProd")
'    AfterLabel.Caption = Param.Text("After")
    BundlesLabel.Caption = Param.Text("Bundles")
    
    'PresetFrame.Caption = Param.Text("PresetQuantity")
    'QuantityTubes.Caption = Param.Text("Tubes")
    'QuantityBundles.Caption = Param.Text("Bundles")
    'PriorityFrame.Caption = Param.Text("Priority")
    'TubesOption.Caption = Param.Text("Tubes")
    'BundlesOption.Caption = Param.Text("Bundles")
    
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
End Sub




'GESTIONE FINE ORDINE
Private Sub FineOrdine_NonStop_Click()
    ModOrder.ModoCambioOrdine = 0
    FineOrdine_NonStop.BackColor = RGB(0, 200, 0)
    FineOrdine_Arresto.BackColor = RGB(220, 220, 220)
    FineOrdine_NuovoOrdine.BackColor = RGB(220, 220, 220)
End Sub

Private Sub FineOrdine_Arresto_Click()
    ModOrder.ModoCambioOrdine = 1
    FineOrdine_NonStop.BackColor = RGB(220, 220, 220)
    FineOrdine_Arresto.BackColor = RGB(0, 200, 0)
    FineOrdine_NuovoOrdine.BackColor = RGB(220, 220, 220)
End Sub

Private Sub FineOrdine_NuovoOrdine_Click()
    ModOrder.ModoCambioOrdine = 2
    FineOrdine_NonStop.BackColor = RGB(220, 220, 220)
    FineOrdine_Arresto.BackColor = RGB(220, 220, 220)
    FineOrdine_NuovoOrdine.BackColor = RGB(0, 200, 0)
End Sub




Private Sub Command1_Click(Index As Integer)

End Sub

Private Sub CommandPgUp_Click()
            RecipeDBGrid.Scroll 0, -10
End Sub

Private Sub CommandPgDn_Click()
            RecipeDBGrid.Scroll 0, 10
End Sub

'Private Sub TubesDisplay_Click()
'    TOUCHNumericPad.Dati = TubesDisplay.Text
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        TubesDisplay.Text = TOUCHNumericPad.Dati
'    End If
'End Sub

Private Sub BundlesEdit_Click()
    TOUCHNumericPad.Dati = OrderBundles.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        OrderBundles.Text = TOUCHNumericPad.Dati
    End If
End Sub

'Private Sub WidthDisplay_CLICK()
'    TOUCHNumericPad.Dati = WidthDisplay.Text
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        WidthDisplay.Text = TOUCHNumericPad.Dati
'    End If
'End Sub

'Private Sub HeightDisplay_CLICK()
'    TOUCHNumericPad.Dati = HeightDisplay.Text
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        HeightDisplay.Text = TOUCHNumericPad.Dati
'    End If
'End Sub

Private Sub ThickDisplay_CLICK()
    TOUCHNumericPad.Dati = ThickDisplay.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        ThickDisplay.Text = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub Label2_Click()

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

Private Sub PrintEdit_Click(Index As Integer)
    Select Case Index
        Case 0
            TOUCHKeyBoard.Dati = PrintEdit(0).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(0).Text = TOUCHKeyBoard.Dati
            End If
        Case 1
            TOUCHKeyBoard.Dati = PrintEdit(1).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(1).Text = TOUCHKeyBoard.Dati
            End If
        Case 2
            TOUCHKeyBoard.Dati = PrintEdit(2).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(2).Text = TOUCHKeyBoard.Dati
            End If
        Case 3
            TOUCHKeyBoard.Dati = PrintEdit(3).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(3).Text = TOUCHKeyBoard.Dati
            End If
        Case 4
            TOUCHKeyBoard.Dati = PrintEdit(4).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(4).Text = TOUCHKeyBoard.Dati
            End If
        Case 5
            TOUCHKeyBoard.Dati = PrintEdit(5).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(5).Text = TOUCHKeyBoard.Dati
            End If
        Case 6
            TOUCHKeyBoard.Dati = PrintEdit(6).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(6).Text = TOUCHKeyBoard.Dati
            End If
        Case 7
            TOUCHKeyBoard.Dati = PrintEdit(7).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(7).Text = TOUCHKeyBoard.Dati
            End If
        Case 8
            TOUCHKeyBoard.Dati = PrintEdit(8).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(8).Text = TOUCHKeyBoard.Dati
            End If
        Case 9
            TOUCHKeyBoard.Dati = PrintEdit(9).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(9).Text = TOUCHKeyBoard.Dati
            End If
        Case 10
            TOUCHKeyBoard.Dati = PrintEdit(10).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(10).Text = TOUCHKeyBoard.Dati
            End If
        Case 11
            TOUCHKeyBoard.Dati = PrintEdit(11).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(11).Text = TOUCHKeyBoard.Dati
            End If
        Case 12
            TOUCHKeyBoard.Dati = PrintEdit(12).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(12).Text = TOUCHKeyBoard.Dati
            End If
        Case 13
            TOUCHKeyBoard.Dati = PrintEdit(13).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(13).Text = TOUCHKeyBoard.Dati
            End If
        Case 14
            TOUCHKeyBoard.Dati = PrintEdit(14).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(14).Text = TOUCHKeyBoard.Dati
            End If
        Case 15
            TOUCHKeyBoard.Dati = PrintEdit(15).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(15).Text = TOUCHKeyBoard.Dati
            End If
        Case 16
            TOUCHKeyBoard.Dati = PrintEdit(16).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(16).Text = TOUCHKeyBoard.Dati
            End If
        Case 17
            TOUCHKeyBoard.Dati = PrintEdit(17).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(17).Text = TOUCHKeyBoard.Dati
            End If
        Case 18
            TOUCHKeyBoard.Dati = PrintEdit(18).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(18).Text = TOUCHKeyBoard.Dati
            End If
        Case 19
            TOUCHKeyBoard.Dati = PrintEdit(19).Text
            TOUCHKeyBoard.Show vbModal
            If TOUCHKeyBoard.DatiConfermati Then
                PrintEdit(19).Text = TOUCHKeyBoard.Dati
            End If
    End Select
End Sub


Private Sub Form_Activate()
    Ok = False
    SpecialBundleFlag = False
    SpecialStrapFlag = False
    
    TitleDisplay.Text = ModOrder.Descrizione
    
    
End Sub

'NON + USATA
' funzione da richiamare prima della Form.show per caricare i dati
Public Sub GetData(Source As OrderClass, OnLineModify As Boolean, TabPosition As Integer)
    
    SpecialBundleFlag = False
    SpecialStrapFlag = False
    ' aggiorna immagine tubi e pacco
    BundleTabUpdate
    TubeDimensioDisplayUpdate
    ' aggiorna immagine reggie
    StrapTabUpdate
    ' aggiorna immagine etichetta e quantità programmata
    LabelTabUpdate
    

    'consenso modifica forma tubo e pacco
    If OnLineModify Then
        BundleShapeFrame.Visible = False
        TubeShapeFrame.Visible = False
    Else
        BundleShapeFrame.Visible = True
        TubeShapeFrame.Visible = True
    End If
    
    SSTab1.Tab = TabPosition

    Ok = False


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Ok = False
End Sub

'***********************************************************
' Funzioni di risposta ai comandi dell'operatore ok e cancel
'***********************************************************
Private Sub CancelCommand_Click()
    Ok = False
    Me.Hide
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
    
    ModOrder.Descrizione = TitleDisplay.Text
'    If CheckData Then
        ' chiude il form
        Ok = True
        Me.Hide
'    Else
'        Ok = False
'    End If
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

Private Sub ItemDisplay_Click()
    TOUCHKeyBoard.Dati = ItemDisplay.Text
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        ItemDisplay.Text = TOUCHKeyBoard.Dati
    End If
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

'***********************************************************************************
' Funzioni di risposta ai comandi dell'operatore in zona etichetta e cambio commessa
'***********************************************************************************

' selezione tipo di cambio commessa
Private Sub ContOption_Click()
    OrderBundles.BackColor = RGB(200, 200, 200)
End Sub
'                ' selezione nome ricetta
'                Private Sub RecipeDBGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'                    ItemDisplay.Text = RecipeData.Recordset.Fields("Item")
'                End Sub

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

