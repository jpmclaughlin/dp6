VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form RecipeModifyForm 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   11400
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   224
      Top             =   10560
      Width           =   2805
   End
   Begin VB.CommandButton ModFileDisegno 
      BackColor       =   &H000000FF&
      Caption         =   "Chiudi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   10560
      Width           =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   10500
   End
   Begin VB.CommandButton CancelCommand 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
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
      Height          =   850
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10560
      Width           =   2300
   End
   Begin VB.CommandButton OkCommand 
      BackColor       =   &H0000FF00&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10560
      Width           =   2300
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
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
            Picture         =   "RecipeModifyForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RecipeModifyForm.frx":0603
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10515
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   18547
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   1058
      BackColor       =   14201263
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Bundle"
      TabPicture(0)   =   "RecipeModifyForm.frx":0C15
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameRicette"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameModFile"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameTubo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameRecipeName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FramePacco"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Strap"
      TabPicture(1)   =   "RecipeModifyForm.frx":0C31
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameRegge"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Facing"
      TabPicture(2)   =   "RecipeModifyForm.frx":0C4D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "Label9(0)"
      Tab(2).Control(2)=   "Label10(0)"
      Tab(2).Control(3)=   "Label9(1)"
      Tab(2).Control(4)=   "Label10(1)"
      Tab(2).Control(5)=   "Label9(2)"
      Tab(2).Control(6)=   "Label10(2)"
      Tab(2).Control(7)=   "Label9(3)"
      Tab(2).Control(8)=   "Label10(3)"
      Tab(2).Control(9)=   "Combo3"
      Tab(2).Control(10)=   "Command3"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "MPS"
      TabPicture(3)   =   "RecipeModifyForm.frx":0C69
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameBypass"
      Tab(3).Control(1)=   "ControlloModMagneti"
      Tab(3).Control(2)=   "ControlloVelTR"
      Tab(3).Control(3)=   "Label15"
      Tab(3).Control(4)=   "Label8"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Ticket"
      TabPicture(4)   =   "RecipeModifyForm.frx":0C85
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label24"
      Tab(4).Control(1)=   "LblCartellini"
      Tab(4).Control(2)=   "Label21"
      Tab(4).Control(3)=   "SelettoreSagome"
      Tab(4).Control(4)=   "Label22"
      Tab(4).Control(5)=   "Label01(1)"
      Tab(4).Control(6)=   "SSTab2"
      Tab(4).Control(7)=   "OptionLingua(0)"
      Tab(4).Control(8)=   "OptionLingua(1)"
      Tab(4).Control(9)=   "OptionLingua(3)"
      Tab(4).Control(10)=   "OptionLingua(4)"
      Tab(4).Control(11)=   "OptionLingua(2)"
      Tab(4).Control(12)=   "OptionLingua(5)"
      Tab(4).ControlCount=   13
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Item code search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   -71790
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   6510
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   -69810
         Style           =   2  'Dropdown List
         TabIndex        =   235
         Top             =   4050
         Width           =   7095
      End
      Begin VB.OptionButton OptionLingua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Special"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   -62400
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   5340
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Frame FrameBypass 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   7725
         Left            =   -70170
         TabIndex        =   182
         Top             =   1020
         Width           =   4395
         Begin VB.CommandButton Command1 
            Caption         =   "Bypass"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   187
            Top             =   3840
            Width           =   3345
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Walkingbeam"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   186
            Top             =   3210
            Width           =   3345
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Verniciatrice"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   1590
            Width           =   3345
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Caricatore"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   210
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   930
            Width           =   3345
         End
         Begin VB.CommandButton CmdPercorso 
            Caption         =   "Cambio percorso"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   5880
            Width           =   3345
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Entrata"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   300
            TabIndex        =   189
            Top             =   270
            Width           =   3210
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "By pass"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   630
            Left            =   360
            TabIndex        =   188
            Top             =   2550
            Width           =   3210
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   7335
            Index           =   3
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   3795
         End
      End
      Begin VB.OptionButton OptionLingua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Francais"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   -62400
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   4560
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.OptionButton OptionLingua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deutsch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   -62430
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   3750
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.OptionButton OptionLingua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Espanol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   3
         Left            =   -62430
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   2970
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.OptionButton OptionLingua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "English"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   1
         Left            =   -62430
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   2190
         Width           =   2355
      End
      Begin VB.OptionButton OptionLingua 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Italiano"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   0
         Left            =   -62430
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   1380
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Frame FrameRegge 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9495
         Left            =   -74805
         TabIndex        =   3
         Top             =   780
         Width           =   14580
         Begin VB.Frame Frame3 
            Height          =   2565
            Left            =   450
            TabIndex        =   215
            Top             =   6930
            Width           =   13485
            Begin VB.ComboBox ComboStrap 
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   660
               Style           =   2  'Dropdown List
               TabIndex        =   216
               Top             =   1290
               Width           =   3075
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
               Left            =   10260
               TabIndex        =   221
               Top             =   930
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Image Selettore 
               Height          =   1155
               Index           =   2
               Left            =   10260
               Top             =   1230
               Visible         =   0   'False
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
               Index           =   0
               Left            =   6810
               TabIndex        =   220
               Top             =   930
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Image Selettore 
               Height          =   1155
               Index           =   3
               Left            =   6810
               Top             =   1230
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Strapping mode"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   510
               Left            =   510
               TabIndex        =   219
               Top             =   570
               Width           =   3270
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Strapping 1 enable"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   510
               Left            =   9240
               TabIndex        =   218
               Top             =   390
               Visible         =   0   'False
               Width           =   3270
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Strapping 2 enable"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   510
               Left            =   5580
               TabIndex        =   217
               Top             =   390
               Visible         =   0   'False
               Width           =   3270
            End
         End
         Begin VB.Frame Frame2 
            Height          =   3165
            Left            =   450
            TabIndex        =   201
            Top             =   3720
            Width           =   13515
            Begin dp6.ControlloUpDown ControlloUpDownRegge 
               Height          =   1005
               Left            =   5535
               TabIndex        =   202
               Top             =   2010
               Width           =   3345
               _ExtentX        =   5900
               _ExtentY        =   1773
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   1
               Left            =   360
               TabIndex        =   214
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   2
               Left            =   1365
               TabIndex        =   213
               Top             =   1140
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   3
               Left            =   2370
               TabIndex        =   212
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   4
               Left            =   3375
               TabIndex        =   211
               Top             =   1140
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   5
               Left            =   4380
               TabIndex        =   210
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   6
               Left            =   5385
               TabIndex        =   209
               Top             =   1140
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   7
               Left            =   6390
               TabIndex        =   208
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   8
               Left            =   7395
               TabIndex        =   207
               Top             =   1140
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   9
               Left            =   8385
               TabIndex        =   206
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "10"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   10
               Left            =   9390
               TabIndex        =   205
               Top             =   1140
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "11"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   11
               Left            =   10395
               TabIndex        =   204
               Top             =   300
               Width           =   1800
            End
            Begin VB.Label TextReggia 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "12"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Index           =   12
               Left            =   11400
               TabIndex        =   203
               Top             =   1140
               Width           =   1800
            End
         End
         Begin dp6.ControlloRegge ControlloRegge1 
            Height          =   1905
            Left            =   840
            TabIndex        =   4
            Top             =   840
            Width           =   12840
            _ExtentX        =   9922
            _ExtentY        =   3254
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   0
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
                  Picture         =   "RecipeModifyForm.frx":0CA1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "RecipeModifyForm.frx":12A4
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Strapping offset [inch]"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   3600
            TabIndex        =   194
            Top             =   3120
            Width           =   3270
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   7200
            TabIndex        =   193
            Top             =   3000
            Width           =   1620
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   7260
            TabIndex        =   192
            Top             =   720
            Width           =   1380
         End
      End
      Begin dp6.ControlloUpDown ControlloModMagneti 
         Height          =   990
         Left            =   -74340
         TabIndex        =   5
         Top             =   2460
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1746
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   9525
         Left            =   -74850
         TabIndex        =   129
         Top             =   900
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   16801
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   882
         BackColor       =   14201263
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Manuali"
         TabPicture(0)   =   "RecipeModifyForm.frx":18B6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Line1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Line1(2)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Line1(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Line1(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Line1(5)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Line1(6)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Line1(7)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Line1(8)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Line1(9)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "campo(1)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "campo(2)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "campo(10)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "campo(9)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "campo(8)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "campo(7)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "campo(6)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "campo(5)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "campo(4)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "campo(3)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Text1(1)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "Text1(2)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Text1(10)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Text1(9)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Text1(8)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Text1(7)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Text1(6)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Text1(5)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Text1(4)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Text1(3)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).ControlCount=   30
         TabCaption(1)   =   "Automatici"
         TabPicture(1)   =   "RecipeModifyForm.frx":18D2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Line1(19)"
         Tab(1).Control(1)=   "Line1(18)"
         Tab(1).Control(2)=   "Line1(17)"
         Tab(1).Control(3)=   "Line1(16)"
         Tab(1).Control(4)=   "Line1(15)"
         Tab(1).Control(5)=   "Line1(14)"
         Tab(1).Control(6)=   "Line1(13)"
         Tab(1).Control(7)=   "Line1(12)"
         Tab(1).Control(8)=   "Line1(11)"
         Tab(1).Control(9)=   "Line1(10)"
         Tab(1).Control(10)=   "campo(13)"
         Tab(1).Control(11)=   "campo(14)"
         Tab(1).Control(12)=   "campo(15)"
         Tab(1).Control(13)=   "campo(16)"
         Tab(1).Control(14)=   "campo(17)"
         Tab(1).Control(15)=   "campo(18)"
         Tab(1).Control(16)=   "campo(19)"
         Tab(1).Control(17)=   "campo(20)"
         Tab(1).Control(18)=   "campo(12)"
         Tab(1).Control(19)=   "campo(11)"
         Tab(1).Control(20)=   "Text1(13)"
         Tab(1).Control(21)=   "Text1(14)"
         Tab(1).Control(22)=   "Text1(15)"
         Tab(1).Control(23)=   "Text1(16)"
         Tab(1).Control(24)=   "Text1(17)"
         Tab(1).Control(25)=   "Text1(18)"
         Tab(1).Control(26)=   "Text1(19)"
         Tab(1).Control(27)=   "Text1(20)"
         Tab(1).Control(28)=   "Text1(12)"
         Tab(1).Control(29)=   "Text1(11)"
         Tab(1).ControlCount=   30
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   3
            Left            =   5280
            TabIndex        =   149
            Top             =   2145
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   4
            Left            =   5310
            TabIndex        =   148
            Top             =   7845
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   5
            Left            =   5280
            TabIndex        =   147
            Top             =   3525
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   6
            Left            =   5280
            TabIndex        =   146
            Top             =   4215
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   7
            Left            =   5280
            TabIndex        =   145
            Top             =   4905
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   8
            Left            =   5280
            TabIndex        =   144
            Top             =   5595
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   9
            Left            =   5280
            TabIndex        =   143
            Top             =   6285
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   10
            Left            =   5280
            TabIndex        =   142
            Top             =   2835
            Width           =   6315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   2
            Left            =   5280
            TabIndex        =   141
            Top             =   1455
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   1
            Left            =   5280
            TabIndex        =   140
            Top             =   750
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   11
            Left            =   -69720
            TabIndex        =   139
            Top             =   750
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   12
            Left            =   -69660
            TabIndex        =   138
            Top             =   9885
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   20
            Left            =   -69720
            TabIndex        =   137
            Top             =   5625
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   19
            Left            =   -69630
            TabIndex        =   136
            Top             =   10125
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   18
            Left            =   -69720
            TabIndex        =   135
            Top             =   6315
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   17
            Left            =   -69720
            TabIndex        =   134
            Top             =   4905
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   16
            Left            =   -69720
            TabIndex        =   133
            Top             =   4215
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   15
            Left            =   -69720
            TabIndex        =   132
            Top             =   3525
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   14
            Left            =   -69720
            TabIndex        =   131
            Top             =   2835
            Width           =   6375
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   13
            Left            =   -69720
            TabIndex        =   130
            Top             =   2145
            Width           =   6375
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   3
            Left            =   270
            TabIndex        =   169
            Top             =   2190
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   4
            Left            =   300
            TabIndex        =   168
            Top             =   7905
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   5
            Left            =   270
            TabIndex        =   167
            Top             =   3615
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   6
            Left            =   270
            TabIndex        =   166
            Top             =   4260
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   7
            Left            =   270
            TabIndex        =   165
            Top             =   4965
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   8
            Left            =   270
            TabIndex        =   164
            Top             =   5655
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   9
            Left            =   270
            TabIndex        =   163
            Top             =   6315
            Width           =   4725
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   10
            Left            =   270
            TabIndex        =   162
            Top             =   2895
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   2
            Left            =   270
            TabIndex        =   161
            Top             =   1515
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   495
            Index           =   1
            Left            =   270
            TabIndex        =   160
            Top             =   870
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Data"
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
            Height          =   525
            Index           =   11
            Left            =   -74760
            TabIndex        =   159
            Top             =   840
            Width           =   4800
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ora"
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
            Height          =   525
            Index           =   12
            Left            =   -74700
            TabIndex        =   158
            Top             =   9945
            Width           =   4800
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "campo 1"
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
            Height          =   525
            Index           =   20
            Left            =   -74760
            TabIndex        =   157
            Top             =   5685
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Spessore tubo"
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
            Height          =   525
            Index           =   19
            Left            =   -74670
            TabIndex        =   156
            Top             =   10185
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Peso pacco"
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
            Height          =   525
            Index           =   18
            Left            =   -74760
            TabIndex        =   155
            Top             =   6405
            Width           =   4770
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Numero tubi"
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
            Height          =   555
            Index           =   17
            Left            =   -74760
            TabIndex        =   154
            Top             =   4965
            Width           =   4755
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Numero pacco"
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
            Height          =   555
            Index           =   16
            Left            =   -74760
            TabIndex        =   153
            Top             =   4245
            Width           =   4770
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lunghezza tubo"
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
            Height          =   525
            Index           =   15
            Left            =   -74760
            TabIndex        =   152
            Top             =   3555
            Width           =   4785
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Dimensione tubo"
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
            Height          =   525
            Index           =   14
            Left            =   -74760
            TabIndex        =   151
            Top             =   2865
            Width           =   4785
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Descrizione"
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
            Height          =   525
            Index           =   13
            Left            =   -74760
            TabIndex        =   150
            Top             =   2205
            Width           =   4800
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   10
            X1              =   -74640
            X2              =   -69240
            Y1              =   2085
            Y2              =   2085
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   11
            X1              =   -74580
            X2              =   -69180
            Y1              =   6225
            Y2              =   6225
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   12
            X1              =   -74580
            X2              =   -69180
            Y1              =   5565
            Y2              =   5565
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   13
            X1              =   -74580
            X2              =   -69180
            Y1              =   4845
            Y2              =   4845
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   14
            X1              =   -74580
            X2              =   -69180
            Y1              =   4185
            Y2              =   4185
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   15
            X1              =   -74640
            X2              =   -69240
            Y1              =   3465
            Y2              =   3465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   16
            X1              =   -74640
            X2              =   -69240
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   17
            X1              =   -74580
            X2              =   -69180
            Y1              =   6945
            Y2              =   6945
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   18
            X1              =   -74580
            X2              =   -69180
            Y1              =   7605
            Y2              =   7605
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   19
            X1              =   -74670
            X2              =   -69270
            Y1              =   1395
            Y2              =   1395
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   9
            X1              =   420
            X2              =   5820
            Y1              =   7605
            Y2              =   7605
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   8
            X1              =   420
            X2              =   5820
            Y1              =   6915
            Y2              =   6915
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   7
            X1              =   360
            X2              =   5760
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   6
            X1              =   360
            X2              =   5760
            Y1              =   3465
            Y2              =   3465
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   420
            X2              =   5820
            Y1              =   4155
            Y2              =   4155
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   4
            X1              =   420
            X2              =   5820
            Y1              =   4845
            Y2              =   4845
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   3
            X1              =   420
            X2              =   5820
            Y1              =   5535
            Y2              =   5535
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   420
            X2              =   5820
            Y1              =   6195
            Y2              =   6195
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   1
            X1              =   360
            X2              =   5760
            Y1              =   2085
            Y2              =   2085
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   0
            X1              =   360
            X2              =   5760
            Y1              =   1395
            Y2              =   1395
         End
      End
      Begin dp6.ControlloUpDown ControlloVelTR 
         Height          =   990
         Left            =   -74370
         TabIndex        =   190
         Top             =   4890
         Visible         =   0   'False
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1746
      End
      Begin VB.Frame FramePacco 
         Caption         =   "Pacco"
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
         Height          =   8370
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   8760
         Begin VB.CommandButton ComModFile 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SPECIALE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   270
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   7140
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Height          =   7845
            Left            =   60
            TabIndex        =   9
            Top             =   450
            Width           =   2325
            Begin VB.ComboBox Combo2 
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
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   200
               Top             =   570
               Width           =   2145
            End
            Begin VB.OptionButton OptionPaccoEsagono 
               BackColor       =   &H0000FF00&
               DownPicture     =   "RecipeModifyForm.frx":18EE
               Height          =   1710
               Left            =   270
               Picture         =   "RecipeModifyForm.frx":B6B0
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   1320
               Width           =   1800
            End
            Begin VB.OptionButton OptionPaccoSQRD 
               BackColor       =   &H0000FF00&
               DownPicture     =   "RecipeModifyForm.frx":15472
               Height          =   1710
               Left            =   240
               Picture         =   "RecipeModifyForm.frx":1F234
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   3090
               Width           =   1800
            End
            Begin VB.OptionButton OptionPaccoSQSQ 
               BackColor       =   &H0000FF00&
               DownPicture     =   "RecipeModifyForm.frx":28FF6
               Height          =   1710
               Left            =   240
               MaskColor       =   &H000000FF&
               Picture         =   "RecipeModifyForm.frx":32DB8
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   4860
               Width           =   1800
            End
         End
         Begin dp6.ControlloUpDownFile OggettoUpDownFile 
            Height          =   3105
            Left            =   7380
            TabIndex        =   13
            Top             =   2160
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   5477
         End
         Begin dp6.ControlloUpDown UpDownFilaBasePacco 
            Height          =   990
            Left            =   3390
            TabIndex        =   14
            Top             =   6540
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   1746
         End
         Begin dp6.ControlloPacco PaccoModRicetta 
            Height          =   5640
            Left            =   3030
            TabIndex        =   15
            Top             =   750
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   9948
         End
         Begin VB.Label LblPesoModOrdine 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "xxxxxx"
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
            Height          =   405
            Left            =   6420
            TabIndex        =   181
            Top             =   7830
            Width           =   2235
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   5490
            TabIndex        =   180
            Top             =   7800
            Width           =   3225
         End
         Begin VB.Label lblModOrdine 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "xxxx"
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
            Height          =   555
            Left            =   3990
            TabIndex        =   178
            Top             =   7830
            Width           =   1365
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "N.Tubi"
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
            Height          =   465
            Left            =   2460
            TabIndex        =   179
            Top             =   7800
            Width           =   2985
         End
      End
      Begin VB.Frame FrameRecipeName 
         Caption         =   "Ricetta"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   120
         TabIndex        =   176
         Top             =   720
         Width           =   15015
         Begin VB.ComboBox ComboStorage 
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
            ItemData        =   "RecipeModifyForm.frx":3CB7A
            Left            =   11850
            List            =   "RecipeModifyForm.frx":3CB7C
            Style           =   2  'Dropdown List
            TabIndex        =   239
            Top             =   660
            Width           =   3075
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SELECT FROM LIST"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   8040
            Style           =   1  'Graphical
            TabIndex        =   238
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Storage destination"
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
            Height          =   1110
            Left            =   11730
            TabIndex        =   240
            Top             =   210
            Width           =   3255
         End
         Begin VB.Image Image1 
            Height          =   990
            Left            =   3690
            Picture         =   "RecipeModifyForm.frx":3CB7E
            Stretch         =   -1  'True
            Top             =   270
            Width           =   1230
         End
         Begin VB.Label LblIDRicetta 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   240
            TabIndex        =   177
            Top             =   480
            Width           =   7695
         End
      End
      Begin VB.Frame FrameTubo 
         Caption         =   "Tubo"
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
         Height          =   8370
         Left            =   8880
         TabIndex        =   16
         Top             =   2040
         Width           =   6270
         Begin VB.Frame Frame4 
            Height          =   1425
            Left            =   3750
            TabIndex        =   241
            Top             =   300
            Width           =   2415
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   660
               Left            =   570
               TabIndex        =   243
               Top             =   660
               Width           =   1380
            End
            Begin VB.Label Label25 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFC0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Weight per Feet"
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
               Height          =   330
               Left            =   120
               TabIndex        =   242
               Top             =   210
               Width           =   2160
            End
         End
         Begin dp6.ControlloPacco TuboModRicetta 
            Height          =   2745
            Left            =   840
            TabIndex        =   18
            Top             =   1500
            Width           =   2280
            _ExtentX        =   5292
            _ExtentY        =   5450
         End
         Begin dp6.ControlloRegge ControlloLunghezza 
            Height          =   1845
            Left            =   210
            TabIndex        =   17
            Top             =   6150
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   3254
         End
         Begin VB.Label DisplayLungIN 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   3240
            TabIndex        =   236
            Top             =   6000
            Width           =   1410
         End
         Begin VB.Image ImgProfilo 
            Height          =   2205
            Left            =   3420
            Top             =   3420
            Width           =   2655
         End
         Begin VB.Label DisplayLunghezza 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   1680
            TabIndex        =   19
            Top             =   6000
            Width           =   1410
         End
         Begin VB.Label DisplaySpessore 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   1140
            TabIndex        =   22
            Top             =   690
            Width           =   1380
         End
         Begin VB.Label DisplayAltezza 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   3420
            TabIndex        =   21
            Top             =   2490
            Width           =   1380
         End
         Begin VB.Label DisplayLarghezza 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   1320
            TabIndex        =   20
            Top             =   4410
            Width           =   1380
         End
      End
      Begin VB.Frame FrameModFile 
         Height          =   9585
         Left            =   120
         TabIndex        =   23
         Top             =   810
         Visible         =   0   'False
         Width           =   15015
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   50
            Left            =   10785
            TabIndex        =   73
            Text            =   "50"
            Top             =   7815
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   49
            Left            =   10785
            TabIndex        =   72
            Text            =   "49"
            Top             =   6968
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   48
            Left            =   10785
            TabIndex        =   71
            Text            =   "48"
            Top             =   6127
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   47
            Left            =   10785
            TabIndex        =   70
            Text            =   "47"
            Top             =   5286
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   46
            Left            =   10785
            TabIndex        =   69
            Text            =   "46"
            Top             =   4445
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   45
            Left            =   10785
            TabIndex        =   68
            Text            =   "45"
            Top             =   3604
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   44
            Left            =   10785
            TabIndex        =   67
            Text            =   "44"
            Top             =   2763
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   43
            Left            =   10785
            TabIndex        =   66
            Text            =   "43"
            Top             =   1922
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   42
            Left            =   10785
            TabIndex        =   65
            Text            =   "42"
            Top             =   1081
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   41
            Left            =   10785
            TabIndex        =   64
            Text            =   "41"
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   1
            Left            =   945
            TabIndex        =   63
            Text            =   "1"
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   2
            Left            =   945
            TabIndex        =   62
            Text            =   "2"
            Top             =   1081
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   3
            Left            =   945
            TabIndex        =   61
            Text            =   "3"
            Top             =   1922
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   4
            Left            =   945
            TabIndex        =   60
            Text            =   "4"
            Top             =   2763
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   5
            Left            =   945
            TabIndex        =   59
            Text            =   "5"
            Top             =   3604
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   6
            Left            =   945
            TabIndex        =   58
            Text            =   "6"
            Top             =   4445
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   7
            Left            =   945
            TabIndex        =   57
            Text            =   "7"
            Top             =   5286
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   8
            Left            =   945
            TabIndex        =   56
            Text            =   "8"
            Top             =   6127
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   9
            Left            =   945
            TabIndex        =   55
            Text            =   "9"
            Top             =   6968
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   10
            Left            =   945
            TabIndex        =   54
            Text            =   "10"
            Top             =   7815
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   11
            Left            =   3405
            TabIndex        =   53
            Text            =   "11"
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   12
            Left            =   3405
            TabIndex        =   52
            Text            =   "12"
            Top             =   1081
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   13
            Left            =   3405
            TabIndex        =   51
            Text            =   "13"
            Top             =   1922
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   14
            Left            =   3405
            TabIndex        =   50
            Text            =   "14"
            Top             =   2763
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   15
            Left            =   3405
            TabIndex        =   49
            Text            =   "15"
            Top             =   3604
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   16
            Left            =   3405
            TabIndex        =   48
            Text            =   "16"
            Top             =   4445
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   17
            Left            =   3405
            TabIndex        =   47
            Text            =   "17"
            Top             =   5286
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   18
            Left            =   3405
            TabIndex        =   46
            Text            =   "18"
            Top             =   6127
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   19
            Left            =   3405
            TabIndex        =   45
            Text            =   "19"
            Top             =   6968
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   20
            Left            =   3405
            TabIndex        =   44
            Text            =   "20"
            Top             =   7815
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   21
            Left            =   5865
            TabIndex        =   43
            Text            =   "21"
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   22
            Left            =   5865
            TabIndex        =   42
            Text            =   "22"
            Top             =   1081
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   23
            Left            =   5865
            TabIndex        =   41
            Text            =   "23"
            Top             =   1922
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   24
            Left            =   5865
            TabIndex        =   40
            Text            =   "24"
            Top             =   2763
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   25
            Left            =   5865
            TabIndex        =   39
            Text            =   "25"
            Top             =   3604
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   26
            Left            =   5865
            TabIndex        =   38
            Text            =   "26"
            Top             =   4445
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   27
            Left            =   5865
            TabIndex        =   37
            Text            =   "27"
            Top             =   5286
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   28
            Left            =   5865
            TabIndex        =   36
            Text            =   "28"
            Top             =   6127
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   29
            Left            =   5865
            TabIndex        =   35
            Text            =   "29"
            Top             =   6968
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   30
            Left            =   5865
            TabIndex        =   34
            Text            =   "30"
            Top             =   7815
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   31
            Left            =   8325
            TabIndex        =   33
            Text            =   "31"
            Top             =   240
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   32
            Left            =   8325
            TabIndex        =   32
            Text            =   "32"
            Top             =   1081
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   33
            Left            =   8325
            TabIndex        =   31
            Text            =   "33"
            Top             =   1922
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   34
            Left            =   8325
            TabIndex        =   30
            Text            =   "34"
            Top             =   2763
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   35
            Left            =   8325
            TabIndex        =   29
            Text            =   "35"
            Top             =   3604
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   36
            Left            =   8325
            TabIndex        =   28
            Text            =   "36"
            Top             =   4445
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   37
            Left            =   8325
            TabIndex        =   27
            Text            =   "37"
            Top             =   5286
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   38
            Left            =   8325
            TabIndex        =   26
            Text            =   "38"
            Top             =   6127
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   39
            Left            =   8325
            TabIndex        =   25
            Text            =   "39"
            Top             =   6968
            Width           =   1500
         End
         Begin VB.TextBox ModFila 
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
            ForeColor       =   &H00FF0000&
            Height          =   780
            Index           =   40
            Left            =   8325
            TabIndex        =   24
            Text            =   "40"
            Top             =   7815
            Width           =   1500
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "49"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   51
            Left            =   10155
            TabIndex        =   123
            Top             =   7035
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "48"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   52
            Left            =   10155
            TabIndex        =   122
            Top             =   6195
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "47"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   53
            Left            =   10155
            TabIndex        =   121
            Top             =   5385
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "46"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   54
            Left            =   10155
            TabIndex        =   120
            Top             =   4530
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "45"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   55
            Left            =   10155
            TabIndex        =   119
            Top             =   3630
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "44"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   56
            Left            =   10155
            TabIndex        =   118
            Top             =   2805
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "43"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   57
            Left            =   10155
            TabIndex        =   117
            Top             =   1920
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "42"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   58
            Left            =   10155
            TabIndex        =   116
            Top             =   1140
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "41"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   59
            Left            =   10155
            TabIndex        =   115
            Top             =   330
            Width           =   600
         End
         Begin VB.Label RowLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "50"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   50
            Left            =   10155
            TabIndex        =   114
            Top             =   7905
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   1
            Left            =   285
            TabIndex        =   113
            Top             =   345
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   2
            Left            =   285
            TabIndex        =   112
            Top             =   1165
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   3
            Left            =   285
            TabIndex        =   111
            Top             =   2000
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   4
            Left            =   285
            TabIndex        =   110
            Top             =   2835
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   5
            Left            =   285
            TabIndex        =   109
            Top             =   3670
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   6
            Left            =   285
            TabIndex        =   108
            Top             =   4505
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   7
            Left            =   285
            TabIndex        =   107
            Top             =   5340
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   8
            Left            =   285
            TabIndex        =   106
            Top             =   6175
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   9
            Left            =   285
            TabIndex        =   105
            Top             =   7010
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   10
            Left            =   285
            TabIndex        =   104
            Top             =   7920
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   11
            Left            =   2760
            TabIndex        =   103
            Top             =   330
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   12
            Left            =   2760
            TabIndex        =   102
            Top             =   1140
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   13
            Left            =   2760
            TabIndex        =   101
            Top             =   1980
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   14
            Left            =   2760
            TabIndex        =   100
            Top             =   2820
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   15
            Left            =   2760
            TabIndex        =   99
            Top             =   3660
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   16
            Left            =   2760
            TabIndex        =   98
            Top             =   4500
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   17
            Left            =   2760
            TabIndex        =   97
            Top             =   5340
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   18
            Left            =   2760
            TabIndex        =   96
            Top             =   6180
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   19
            Left            =   2760
            TabIndex        =   95
            Top             =   7020
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   20
            Left            =   2760
            TabIndex        =   94
            Top             =   7905
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   21
            Left            =   5235
            TabIndex        =   93
            Top             =   330
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   22
            Left            =   5235
            TabIndex        =   92
            Top             =   1140
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   23
            Left            =   5235
            TabIndex        =   91
            Top             =   1980
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "24"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   24
            Left            =   5235
            TabIndex        =   90
            Top             =   2835
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "25"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   25
            Left            =   5235
            TabIndex        =   89
            Top             =   3690
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Height          =   600
            Index           =   26
            Left            =   5235
            TabIndex        =   88
            Top             =   4530
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "27"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   27
            Left            =   5235
            TabIndex        =   87
            Top             =   5385
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "28"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   28
            Left            =   5235
            TabIndex        =   86
            Top             =   6240
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "29"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   29
            Left            =   5235
            TabIndex        =   85
            Top             =   7080
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   30
            Left            =   5235
            TabIndex        =   84
            Top             =   7905
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "31"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   31
            Left            =   7650
            TabIndex        =   83
            Top             =   345
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "32"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   32
            Left            =   7650
            TabIndex        =   82
            Top             =   1230
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   33
            Left            =   7650
            TabIndex        =   81
            Top             =   2040
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "34"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   34
            Left            =   7650
            TabIndex        =   80
            Top             =   2790
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "35"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   35
            Left            =   7650
            TabIndex        =   79
            Top             =   3645
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "36"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   36
            Left            =   7650
            TabIndex        =   78
            Top             =   4590
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "37"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   37
            Left            =   7650
            TabIndex        =   77
            Top             =   5490
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "38"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   38
            Left            =   7650
            TabIndex        =   76
            Top             =   6300
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "39"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   39
            Left            =   7650
            TabIndex        =   75
            Top             =   7170
            Width           =   600
         End
         Begin VB.Label LabelFila 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "40"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   0
            Left            =   7650
            TabIndex        =   74
            Top             =   7920
            Width           =   600
         End
      End
      Begin VB.Frame FrameRicette 
         Caption         =   "Ordine"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   120
         TabIndex        =   124
         Top             =   720
         Width           =   14955
         Begin VB.TextBox DisplayPacchiOrdine 
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
            Height          =   690
            Left            =   13140
            TabIndex        =   127
            Text            =   "100"
            Top             =   480
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   480
            Width           =   3075
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pacchi"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   11640
            TabIndex        =   175
            Top             =   720
            Width           =   1470
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dopo"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   11580
            TabIndex        =   128
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label DisplayDescrizione 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   240
            TabIndex        =   125
            Top             =   480
            Width           =   7695
         End
         Begin VB.Shape Shape13 
            BorderColor     =   &H00000000&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Left            =   60
            Shape           =   4  'Rounded Rectangle
            Top             =   420
            Width           =   14775
         End
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   3
         Left            =   -69810
         TabIndex        =   233
         Top             =   5310
         Width           =   7050
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "itemcode"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         Left            =   -72630
         TabIndex        =   232
         Top             =   5430
         Width           =   3270
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   2
         Left            =   -71100
         TabIndex        =   231
         Top             =   1680
         Visible         =   0   'False
         Width           =   6150
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pieces"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   2
         Left            =   -74700
         TabIndex        =   230
         Top             =   1800
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   1
         Left            =   -71100
         TabIndex        =   229
         Top             =   780
         Visible         =   0   'False
         Width           =   6150
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         Left            =   -74700
         TabIndex        =   228
         Top             =   900
         Visible         =   0   'False
         Width           =   3270
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   0
         Left            =   -69810
         TabIndex        =   227
         Top             =   3330
         Width           =   7080
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
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
         Left            =   -72630
         TabIndex        =   226
         Top             =   3450
         Width           =   3270
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Metric          Inch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   -62430
         TabIndex        =   198
         Top             =   6960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit"
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
         Height          =   390
         Left            =   -62250
         TabIndex        =   197
         Top             =   6240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image SelettoreSagome 
         Height          =   1275
         Left            =   -61800
         Top             =   7290
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "n. cartellini"
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
         Height          =   585
         Left            =   -62250
         TabIndex        =   196
         Top             =   8940
         Width           =   2175
      End
      Begin VB.Label LblCartellini 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
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
         Height          =   585
         Left            =   -62250
         TabIndex        =   195
         Top             =   9510
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Up conveyor speed (%)"
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
         Left            =   -74550
         TabIndex        =   191
         Top             =   4200
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vel. magneti (%)"
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
         Left            =   -74460
         TabIndex        =   6
         Top             =   1680
         Width           =   3540
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2115
         Left            =   -62250
         TabIndex        =   222
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Height          =   6555
         Left            =   -72480
         TabIndex        =   234
         Top             =   2400
         Width           =   9915
      End
   End
   Begin VB.Label Label23 
      Caption         =   "Label4"
      Height          =   765
      Left            =   0
      TabIndex        =   199
      Top             =   0
      Width           =   2985
   End
End
Attribute VB_Name = "RecipeModifyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

'*************************************************
' Costanti e variabili per gestione input dati
'**************************************************
' Ok è public perchè può essere esaminata per verificare
' se i dati sono stati confermati
Public OK As Boolean
' flag protezione modifica  dati
Private FromOperator As Boolean
' memoria modifica pacco speciale in corso
Private SpecialBundleFlag As Boolean
Private SpecialStrapFlag As Boolean
Private LinguaSelezionata As Byte
Private WeightModified As String

Public TicketFrefresh As Boolean
Public ModZona As Boolean
Public StrapEnable As Boolean
Public BundleEnable As Boolean
'==============================================================================
' FUNZIONE AGGIORNAMENTO pagina
'==============================================================================

Sub Update()
    Dim i As Integer
       
    ComboStorage.Visible = Param.GetBit("Par229_STORAGE_DEST")
    Label13.Visible = Param.GetBit("Par229_STORAGE_DEST")
    Cartellino.Lingua = Ing  'forzatura cartellino inglese
    If Cartellino.Lingua > 0 Then
       OptionLingua(Cartellino.Lingua).value = True
    Else
       OptionLingua(1).value = True
    End If
    For i = 1 To 20
        campo(i).Visible = Text1(i).Visible
    Next
End Sub

Private Sub CmdPercorso_Click()
   
   frmAvvisi.AvvisoBypass = True
   frmAvvisi.Show vbModal

   TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
   TechPasswordForm.Show vbModal
   If TechPasswordForm.LoginSucceeded = False Then Exit Sub
   Command1(0).Enabled = True
   Command1(1).Enabled = True
   Command1(2).Enabled = True
   Command1(3).Enabled = True
   Unload TechPasswordForm
End Sub

Private Sub Combo1_Click()
   Select Case Combo1.ListIndex
      Case 0
               ModOrder.ModoCambioOrdine = 0
               Label1.Visible = False
               Label2.Visible = False
               DisplayPacchiOrdine.Visible = False
      Case 1
               ModOrder.ModoCambioOrdine = 1
               Label1.Visible = True
               Label2.Visible = True
               DisplayPacchiOrdine.Visible = True
      Case 2
                ModOrder.ModoCambioOrdine = 2
                Label1.Visible = True
                Label2.Visible = True
                DisplayPacchiOrdine.Visible = True
     End Select
End Sub

Private Sub Combo2_Click()
    ImgProfilo.Visible = Combo2.Visible And Combo2.ListIndex > 0
    If Combo2.ListIndex > 0 Then ImgProfilo.Picture = LoadPicture("..\Bitmap\" & NomeProfili(Combo2.ListIndex) & ".gif")
    Ricetta.Profilo = Combo2.ListIndex
    Call ControlloDatiPacco
    RefreshDati
    DisegnoPaccoTuboRefresh
End Sub

Private Sub Combo3_Click()
   Ricetta.Grade = Combo3.Text
   RefreshDati
End Sub

Private Sub ComboStorage_Click()
    Select Case ComboStorage.ListIndex
'    Case 0
'        DB426.Word(70) = 0
    Case 1
        Ricetta.Destination = 1
    Case 2
        Ricetta.Destination = 2
    Case 3
        Ricetta.Destination = 3
    Case Else
        Ricetta.Destination = 0
    End Select
End Sub

Private Sub ComboStrap_Click()
   Ricetta.TipoCalcRegge = ComboStrap.ListIndex
   Label20.Visible = ComboStrap.ListIndex
   Label5.Visible = ComboStrap.ListIndex
   Ricetta.CalcoloQuoteRegge
   Call RefreshRegge
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    Case 0
       Ricetta.Bypass0 = True
       Ricetta.Bypass1 = False
       Command1(0).BackColor = &HFF00&
       Command1(1).BackColor = &H8000000F
    Case 1
       Ricetta.Bypass0 = False
       Ricetta.Bypass1 = True
       Command1(0).BackColor = &H8000000F
       Command1(1).BackColor = &HFF00&
    Case 2
       Ricetta.Bypass2 = True
       Ricetta.Bypass3 = False
       Command1(2).BackColor = &HFF00&
       Command1(3).BackColor = &H8000000F
    Case 3
       Ricetta.Bypass2 = False
       Ricetta.Bypass3 = True
       Command1(2).BackColor = &H8000000F
       Command1(3).BackColor = &HFF00&
    End Select
End Sub

Private Sub Command2_Click()
   frmStampa.RefreshVar
   frmStampa.Command1.Visible = True
   frmStampa.Show vbModal
   frmStampa.Command1.Visible = False
End Sub

Private Sub Command3_Click()
    Dim rst As New ADODB.Recordset
    
    Debug.Print "SELECT DISTINCT Size  FROM Recipes_itemesCode WHERE Size like '%" & Trim(Replace(Left(Ricetta.IDRicetta, 9), "X", " X ")) & "%'"
    
    If rst.State = adStateOpen Then rst.Close
    rst.Open "SELECT DISTINCT * FROM Recipes_itemesCode", Connessione, adOpenKeyset, adLockReadOnly
    rst.MoveFirst
    rst.Find "Size='" & Trim(Replace(Replace(Left(Ricetta.IDRicetta, 9), "X", " X "), "SCH", "SCH ")) & "'"
    'rst.Open "SELECT DISTINCT * FROM Recipes_itemesCode WHERE Size='" & Trim(Replace(Left(Ricetta.IDRicetta, 9), "X", " X ")) & "'", Connessione, adOpenKeyset, adLockOptimistic, adCmdText
    If rst.EOF Then
       Label10(3) = "Not_Found"
    Else
       Label10(3) = rst.Fields("item_code")
       rst.Close
    End If
    Set rst = Nothing
    
End Sub

Private Sub Command4_Click()
   DialogSerials.Show vbModal
   If Trim(DialogSerials.SerialNumber) <> "" And Trim(DialogSerials.RecipeName) <> "" Then
      If OrderModifyForm.NewOrModify Then Ricetta.IDRicetta = Trim(DialogSerials.RecipeName) & "X" & Conv_UM.Conversione(Ricetta.TuboLunghezza, UM.mt, UM.ft, 0)
      Ricetta.Itemcode = Trim(DialogSerials.SerialNumber) & Conv_UM.Conversione(Ricetta.TuboLunghezza, UM.mt, UM.ft, 0)
      Ricetta.Grade = Trim(DialogSerials.GradeNumber)
      RefreshDati
   End If
   Unload DialogSerials
End Sub

Private Sub DisplayDescrizione_Click()
    TOUCHKeyBoard.Dati = ModOrder.Descrizione
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        DisplayDescrizione.caption = TOUCHKeyBoard.Dati
        ModOrder.Descrizione = TOUCHKeyBoard.Dati
    End If
End Sub

Private Sub DisplayLungIN_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 11.9
    TOUCHNumericPad.Dati = DisplayLungIN.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
       Ricetta.TuboLunghezza = DisplayLunghezza * 12 * 25.4 / 1000 + TOUCHNumericPad.Dati * 25.4 / 1000
       Ricetta.ControlloDatiPacco
       Ricetta.CalcoloQuoteRegge
       RefreshDati
    End If
End Sub

Private Sub DisplayPacchiOrdine_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1
    TOUCHNumericPad.ValoreMax = 999
    TOUCHNumericPad.Dati = DisplayPacchiOrdine.Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        DisplayPacchiOrdine.Text = TOUCHNumericPad.Dati
        ModOrder.PresetPacchi = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub Form_Load()
    
    With OrdersForm.AdoOrdini
           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
           .CommandType = adCmdTable
           .RecordSource = "_Grade_list"
           .Refresh
           .Recordset.MoveFirst
             
           While .Recordset.EOF = False
               Combo3.AddItem .Recordset.Fields("Grade")
               .Recordset.MoveNext
           Wend
     End With

'    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
 '   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    TicketFrefresh = False
    With rs
        .Open "SELECT * FROM Profili", Connessione, , adLockReadOnly, adCmdText
        Combo2.Clear
        .MoveNext
        Combo2.AddItem .Fields("Alias")
        .MoveNext
        Combo2.AddItem .Fields("Alias")
        NomeProfili(1) = .Fields("Nome")
        .MoveNext
        Combo2.AddItem .Fields("Alias")
        NomeProfili(2) = .Fields("Nome")
'        .MoveNext
'        Combo2.AddItem .Fields("Alias")
'        NomeProfili(3) = .Fields("Nome")
        .Close
        Set .ActiveConnection = Nothing
    End With
    Set rs = Nothing
 '   Set cn = Nothing

    ScritteMultilingua
   ComboStrap.Clear
   ComboStrap.AddItem Param.Text("ReggStandard")
 '  ComboStrap.AddItem "Left fix"
   ComboStrap.AddItem Param.Text("ReggCentre")
   ComboStrap.AddItem "Double strap"
End Sub

'===============================================================================
'--------------------- INIZIO FUNZIONI PAGINA
'===============================================================================

Private Sub Form_Activate()
    Dim IndiceRiga As Integer
    Dim i As Integer
    
    
    ModFileDisegno.Visible = FrameModFile.Visible
    LblCartellini = Param.GetNumber("Par221_NumeroCartellini")
    SelettoreSagome.Picture = ImageList1.ListImages(1 + Abs(CInt(ModOrder.TicketUnit))).Picture
    OptionPaccoSQSQ.Visible = Param.GetBit("Par224_TuboSQ")
    OptionPaccoEsagono.Visible = Param.GetBit("Par180_PaccoHex")
    'attiva modifica lunghezza in zona reggiatura
  '  Shape21.Visible = ModZona
    Label19.Visible = ModZona
    Label5.caption = Conv_UM.Conversione(Param.GetNumber("Par223_OffSetRegg") / 1000, UM.mt, UM.inch, 3)
    If ModZona Then Label19.caption = Conv_UM.Conversione(Ricetta.TuboLunghezza, UM.mt, UM.inch, 1) * 1000
    '-----------
    ComboStrap.ListIndex = Ricetta.TipoCalcRegge
    OK = False
    FrameBypass.Visible = Param.GetBit("Par212_AttivaGestioneBypass")
    Image1.Visible = False
    Timer1.Enabled = False
    If Ricetta.Bypass0 = True Then
      Command1(0).BackColor = &HFF00&
    End If
    If Ricetta.Bypass1 = True Then
      Command1(1).BackColor = &HFF00&
    End If
    If Ricetta.Bypass2 = True Then
      Command1(2).BackColor = &HFF00&
    End If
    If Ricetta.Bypass3 = True Then
      Command1(3).BackColor = &HFF00&
    End If
      
    If Not Not Param.GetBit("Par220_AttivaGestioneRicette") And Not (frmKernel.PaginaCorrente = PagPacco) Then
        FrameRicette.Visible = False
        FrameRecipeName.Visible = True
    Else
        FrameRicette.Visible = True
        FrameRecipeName.Visible = False
    End If
    
    FrameModFile.ZOrder 0
    If ModOrder.Descrizione = "" Then
      ModOrder.Descrizione = "---"
    End If
    
    SSTab1.Width = Width
    LblIDRicetta.caption = Ricetta.IDRicetta
    Combo1.Text = Combo1.List(ModOrder.ModoCambioOrdine)
    DisplayDescrizione.caption = ModOrder.Descrizione
    DisplayPacchiOrdine.Text = ModOrder.PresetPacchi
    Label20.Visible = ComboStrap.ListIndex
    Label5.Visible = ComboStrap.ListIndex
   
     If ModOrder.ModoCambioOrdine > 0 Then
        DisplayPacchiOrdine.Visible = True
     End If
               
    ' abilitazione cartelle modifica
    
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = frmKernel.PaginaCorrente = PagPacco Or frmKernel.PaginaCorrente = PagOrdini
    SSTab1.TabEnabled(3) = False
    SSTab1.TabEnabled(4) = False
    
    PaccoModRicetta.ColoreSfondo = &H8000000F
    TuboModRicetta.ColoreSfondo = &H8000000F
    
    SSTab1.TabEnabled(1) = StrapEnable
    SSTab1.TabEnabled(0) = BundleEnable
    'SSTab1.TabEnabled(4) = PrinterInstall And (Param.GetBit("Par220_AttivaGestioneRicette") = False)
    
    SSTab1.TabEnabled(4) = PrinterInstall And SSTab1.TabEnabled(0) = False And frmKernel.PaginaCorrente <> PagRegge
    
    If Param.GetBit("Par214_MPS") And (frmKernel.PaginaCorrente = PagPacco Or frmKernel.PaginaCorrente = PagOrdini) Then SSTab1.TabEnabled(3) = True
   ' If Param.GetBit("Par213_Fasciatura") And StrapEnable Then SSTab1.TabEnabled(2) = True
    Command2.Visible = SSTab1.TabEnabled(4)
    ' in modifica il nome non può venire cambiato
    LblIDRicetta.Enabled = OrderModifyForm.NewOrModify
    'oggetto modifica base pacco
    UpDownFilaBasePacco.Step = 1
    UpDownFilaBasePacco.LimMax = 100
    UpDownFilaBasePacco.LimMin = 1
    'oggetto modifica file pacco
    OggettoUpDownFile.Step = 1
    OggettoUpDownFile.LimMax = MAX_ROWS
    OggettoUpDownFile.LimMin = 1
    'oggetto modifica regge
    ControlloUpDownRegge.Step = 1
    ControlloUpDownRegge.LimMax = MAX_STRAPS
    ControlloUpDownRegge.LimMin = 1
    'Disegno lunghezza
    DisegnoLunghezzaRefresh
    'oggetto modifica velocità MPS
    ControlloVelTR.Step = 1
    ControlloVelTR.LimMin = 20
    ControlloVelTR.LimMax = 80
    ControlloVelTR.Decimali = 0
    
    ControlloModMagneti.Step = 10
    ControlloModMagneti.LimMax = 100
    ControlloModMagneti.LimMin = 20
    ControlloModMagneti.Decimali = 0
'    ControlloOverrideVR1.Step = 10
'    ControlloOverrideVR1.LimMax = 100
'    ControlloOverrideVR1.LimMin = 20
'    ControlloOverrideVR1.Decimali = 0
'    ControlloOverrideVR2.Step = 10
'    ControlloOverrideVR2.LimMax = 100
'    ControlloOverrideVR2.LimMin = 20
'    ControlloOverrideVR2.Decimali = 0
'    ControlloMonobeam1.Step = 10
'    ControlloMonobeam1.LimMax = 100
'    ControlloMonobeam1.LimMin = 20
'    ControlloMonobeam1.Decimali = 0
'    ControlloMonobeam2.Step = 10
'    ControlloMonobeam2.LimMax = 100
'    ControlloMonobeam2.LimMin = 20
'    ControlloMonobeam2.Decimali = 0
'    ControlloMonobeam3.Step = 10
'    ControlloMonobeam3.LimMax = 100
'    ControlloMonobeam3.LimMin = 20
'    ControlloMonobeam3.Decimali = 0
    
    ComboStorage.Clear
    ComboStorage.AddItem "Auto"
    ComboStorage.AddItem "Storage 1"
    ComboStorage.AddItem "Storage 2"
    ComboStorage.AddItem "Storage 3"
    ComboStorage.ListIndex = Val(Ricetta.Destination)
  
    'rinfresco selettori reggiatrici
    
    Selettore(2).Picture = ImageList1.ListImages(1 + Abs(CInt(Ricetta.Regg1))).Picture
    Selettore(3).Picture = ImageList1.ListImages(1 + Abs(CInt(Ricetta.Regg2))).Picture
    
    'RINFRESCO PAGINA
    
    Call RefreshDati
    Call RefreshRegge

    FromOperator = True
    
    Combo2.ListIndex = Ricetta.Profilo
    If Ricetta.TipoPacco = Pacco.Quadro And Ricetta.TipoTubo = Tubo.Quadro And Param.GetBit("Par201_AbilitazioneProfili") Then
      Combo2.Visible = Param.GetBit("Par201_AbilitazioneProfili")
    Else
      Combo2.Visible = False
    End If
    ImgProfilo.Visible = Combo2.Visible And Combo2.ListIndex > 0
    If Combo2.ListIndex > 0 Then ImgProfilo.Picture = LoadPicture("..\Bitmap\" & NomeProfili(Combo2.ListIndex) & ".gif")
    '=================================================================
    ' aggiornamento cartellino
    Update
    
    If TicketFrefresh Then Exit Sub
    
    WeightToPrint = Ricetta.Weight
    
    For i = 1 To 20
      Text1(i).Visible = Cartellino.FissoVisibile(i)
      campo(i).Visible = Text1(i).Visible
      campo(i) = Cartellino.Fisso(i)
      If i < 11 Then
         Text1(i) = Cartellino.CampoManuale(i)
      Else
         Text1(i) = Cartellino.CampoAuto(i - 10)
      End If
    Next
    TicketFrefresh = True
End Sub

Private Sub RefreshDati()
    If Ricetta.TipoTubo = Tubo.Tondo Then
        Ricetta.TuboAltezza = Ricetta.TuboLarghezza
        DisplayAltezza.Visible = False
    Else
        DisplayAltezza.Visible = True
    End If
        
    PaccoModRicetta.Left = 3050
    PaccoModRicetta.Width = 4200
    
    If Ricetta.TipoPacco = Pacco.Esagono Then
        PaccoModRicetta.Left = 3800
        PaccoModRicetta.Width = 4200
        OptionPaccoEsagono.value = True
        OggettoUpDownFile.Visible = False
     Else
        PaccoModRicetta.Left = 3050
        PaccoModRicetta.Width = 4200
        OggettoUpDownFile.Visible = True
        OggettoUpDownFile.Top = PaccoModRicetta.Top + (PaccoModRicetta.Width / 2) - (OggettoUpDownFile.Width / 2)
        If Ricetta.TipoTubo = Tubo.Tondo Then
            OptionPaccoSQRD.value = True
        Else
            OptionPaccoSQSQ.value = True
        End If
    End If
    UpDownFilaBasePacco.Left = PaccoModRicetta.Left + (PaccoModRicetta.Width / 2) - (UpDownFilaBasePacco.Width / 2) - 350
    UpDownFilaBasePacco.value = Ricetta.TubiFila(1)
    UpDownFilaBasePacco.Refresh
    
    If ControlloVelTR.Occupato = False Then
       ControlloVelTR.value = Ricetta.VelTR
       ControlloVelTR.Refresh
    End If
    
    If ControlloModMagneti.Occupato = False Then
       ControlloModMagneti.value = Ricetta.VelMPS
       ControlloModMagneti.Refresh
    End If
    
'    If ControlloOverrideVR1.Occupato = False Then
'       ControlloOverrideVR1.value = Ricetta.VelVR1
'       ControlloOverrideVR1.Refresh
'    End If
'
'    If ControlloOverrideVR2.Occupato = False Then
'       ControlloOverrideVR2.value = Ricetta.VelVR2
'       ControlloOverrideVR2.Refresh
'    End If
'
'    If ControlloMonobeam1.Occupato = False Then
'       ControlloMonobeam1.value = Ricetta.VelMB1
'       ControlloMonobeam1.Refresh
'    End If
'
'    If ControlloMonobeam2.Occupato = False Then
'       ControlloMonobeam2.value = Ricetta.VelMB2
'       ControlloMonobeam2.Refresh
'    End If
'
'    If ControlloMonobeam3.Occupato = False Then
'       ControlloMonobeam3.value = Ricetta.VelMB3
'       ControlloMonobeam3.Refresh
'    End If
    OggettoUpDownFile.value = Ricetta.NumeroFile
    OggettoUpDownFile.Refresh
    DisegnoPaccoTuboRefresh
    If Ricetta.TipoTubo = Tubo.Tondo Then
        DisplaySpessore.Left = 2800
        DisplayLarghezza.Left = 2300
    Else
        DisplaySpessore.Left = 1600
        DisplayLarghezza.Left = 1100
    End If

    DisplayLarghezza.caption = Conv_UM.Conversione(Ricetta.TuboLarghezza, UM.mt, UM.inch, 2)
    DisplayAltezza.caption = Conv_UM.Conversione(Ricetta.TuboAltezza, UM.mt, UM.inch, 2)
    DisplaySpessore.caption = Conv_UM.Conversione(Ricetta.TuboSpessore, UM.mt, UM.inch, 3)
    Dim UnitTemp As Single
    
    UnitTemp = Ricetta.TuboLunghezza * 1000 / 25.4 / 12
    DisplayLunghezza.caption = Fix(UnitTemp)
    DisplayLungIN.caption = Conv_UM.Conversione(UnitTemp - Fix(UnitTemp), UM.ft, UM.inch, 1)
'    Label12 = Conv_UM.Conversione(Ricetta.TuboLunghezza, UM.mt, UM.ft, 0)
    Label10(0) = Ricetta.Grade
    Label10(1) = Ricetta.Weight
    Label10(2) = Ricetta.Pieces
    Label10(3) = Ricetta.Itemcode
    
    ControlloLunghezza.PaccoLunghezza = Ricetta.TuboLunghezza
    
    Call RefreshRegge
    Call DisegnoLunghezzaRefresh
End Sub

Private Sub RefreshRegge()
    ControlloUpDownRegge.value = Ricetta.NumeroRegge
    ControlloUpDownRegge.Refresh
    Dim i As Integer
    For i = 1 To MAX_STRAPS
        If i > Ricetta.NumeroRegge Then
            TextReggia(i).Visible = False
        Else
            TextReggia(i).Visible = True
        End If
        TextReggia(i).caption = Conv_UM.Conversione(Ricetta.QuotaReggia(i), UM.mt, UM.inch, 2)
        ControlloRegge1.QuotaReggia(i) = Ricetta.QuotaReggia(i)
    Next
     If ModZona = False Then
        ControlloRegge1.VisualizzaLabelLunghezza = True
     Else
        ControlloRegge1.VisualizzaLabelLunghezza = False
     End If
    ControlloRegge1.PaccoLunghezza = Ricetta.TuboLunghezza
    ControlloRegge1.Refresh
End Sub

Private Sub DisegnoLunghezzaRefresh()
    ControlloLunghezza.VisualizzaLabelLunghezza = False
    ControlloLunghezza.VisualizzaQuote = False
    ControlloLunghezza.Refresh
End Sub

Private Sub Label10_Click(Index As Integer)
    Select Case Index
    Case 0
       TOUCHKeyBoard.Dati = Ricetta.Grade
    Case 1
       TOUCHKeyBoard.Dati = Ricetta.Weight
    Case 2
       TOUCHKeyBoard.Dati = Ricetta.Pieces
    Case 3
       TOUCHKeyBoard.Dati = Ricetta.Itemcode
    End Select
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        Select Case Index
        Case 0
           Ricetta.Grade = TOUCHKeyBoard.Dati
        Case 1
           Ricetta.Weight = TOUCHKeyBoard.Dati
        Case 2
           Ricetta.Pieces = TOUCHKeyBoard.Dati
        Case 3
           Ricetta.Itemcode = TOUCHKeyBoard.Dati
        End Select
        Label10(Index) = TOUCHKeyBoard.Dati
    End If
End Sub

Private Sub Label14_Click()
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = 0 'Conv_UM.Conversione(Param.GetNumber("Par004_Tubo_AltezzaMin"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.ValoreMax = 10000 'Conv_UM.Conversione(Param.GetNumber("Par003_Tubo_AltezzaMax"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.Dati = Label14.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.WeightPerFeet = TOUCHNumericPad.Dati
        Ricetta.ControlloDatiPacco
        RefreshDati
    End If
End Sub

Private Sub Label19_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par008_Tubo_LunghezzaMin"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par007_Tubo_LunghezzaMax"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.Dati = Label19.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.TuboLunghezza = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.inch, UM.mt, 3)
        Label19.caption = TOUCHNumericPad.Dati
        Ricetta.ControlloDatiPacco
        Ricetta.CalcoloQuoteRegge
        RefreshDati
    End If
End Sub

Private Sub Label5_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 100
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par007_Tubo_LunghezzaMax"), UM.inch, UM.mt, 3)
    TOUCHNumericPad.Dati = Val(Label5.caption)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Param.SetNumber "Par223_OffSetRegg", TOUCHNumericPad.Dati
            Label5 = TOUCHNumericPad.Dati
        On Error GoTo 0
    End If
End Sub

Private Sub LblCartellini_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1
    TOUCHNumericPad.ValoreMax = 5
    TOUCHNumericPad.Dati = Val(LblCartellini.caption)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Param.SetNumber "Par221_NumeroCartellini", TOUCHNumericPad.Dati
            LblCartellini = TOUCHNumericPad.Dati
        On Error GoTo 0
    End If
End Sub

Private Sub LblIDRicetta_Click()
  TOUCHKeyBoard.Dati = Ricetta.IDRicetta
  TOUCHKeyBoard.Show vbModal
  If TOUCHKeyBoard.DatiConfermati Then
     LblIDRicetta.caption = TOUCHKeyBoard.Dati
     Ricetta.IDRicetta = TOUCHKeyBoard.Dati
  End If
End Sub

'***********************************************************
' Funzioni di risposta ai comandi dell'operatore ok e cancel
'***********************************************************
Private Sub OkCommand_Click()
   Dim i As Integer
   Dim Trovata As Boolean
       
 ModZona = False
 Trovata = False
 If Param.GetBit("Par220_AttivaGestioneRicette") And OrderModifyForm.NewOrModify Then
    With OrdersForm.AdoOrdini
       .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
       .CommandType = adCmdTable
       .RecordSource = "Recipes"
       .Refresh
       .Recordset.MoveFirst
       
       For i = 0 To .Recordset.RecordCount - 1
          If Ricetta.IDRicetta = "????????????" Or .Recordset.Fields("ID") = Ricetta.IDRicetta Then
             Trovata = True: Exit For
          End If
          .Recordset.MoveNext
       Next
       Set .Recordset.ActiveConnection = Nothing
    End With
 End If

   If Trovata = False Then
        OK = True
        OrderModifyForm.PulsantePremuto = True
        ModOrder.LinguaCartellino = LinguaSelezionata
        Cartellino.Lingua = LinguaSelezionata
        For i = 1 To 20
           If i < 11 Then
              ModOrder.CampoManuale(i) = Text1(i).Text
              Cartellino.CampoManuale(i) = Text1(i).Text
           Else
              If SSTab2.TabEnabled(1) = True Then
                 Cartellino.CampoManuale(i) = Text1(i).Text
              End If
           End If
        Next
        
        Cartellino.Scrive_file_dati
        FrameModFile.Visible = False
        FromOperator = False
        While BundleForm.DisegnoTubo.InDisegno = True Or BundleForm.DisegnoPacco.InDisegno = True
           DoEvents
        Wend
        Me.Hide
   Else
      Timer1.Enabled = True
   End If
End Sub

Private Sub CancelCommand_Click()
    OK = False
    ModZona = False
    FrameModFile.Visible = False
    FromOperator = False
    Me.Hide
End Sub

Private Sub OptionLingua_Click(Index As Integer)
    Dim i As Integer

    LinguaSelezionata = Index
    With frmStampa
      '.LoadTicketFile "..\Target\Tickets\MairTicket_" & Format(LinguaSelezionata, "00") & ".TKT"
      .LoadFixTexts
      .LengFixTextRefresh (LinguaSelezionata)
      .FixVisibleRefresh
    End With
    ' aggiorna pagina
    For i = 1 To 20
      Text1(i).Visible = Cartellino.FissoVisibile(i)
      campo(i).Visible = Text1(i).Visible
      campo(i) = Cartellino.Fisso(i)
      If i < 11 Then
         Text1(i) = Cartellino.CampoManuale(i)
      Else
         Text1(i) = Cartellino.CampoAuto(i - 10)
      End If
    Next
End Sub

' OPTION PACCO-TUBO
Private Sub OptionPaccoEsagono_Click()
    Ricetta.TipoPacco = Pacco.Esagono
    Ricetta.TipoTubo = Tubo.Tondo
    Ricetta.Profilo = 0
    Call ControlloDatiPacco
    If FromOperator Then Ricetta.CostruisciPaccoRegolare
    RefreshDati
    DisegnoPaccoTuboRefresh
    Combo2.Visible = False
    ImgProfilo.Visible = Combo2.Visible And Combo2.ListIndex > 0
  End Sub
Private Sub OptionPaccoSQRD_Click()
    If Ricetta.TipoPacco = 1 Then Ricetta.NumeroFile = Ricetta.TubiFila(1)
    Ricetta.TipoPacco = Pacco.Quadro
    Ricetta.TipoTubo = Tubo.Tondo
    Ricetta.Profilo = 0
    Call ControlloDatiPacco
    If FromOperator Then Ricetta.CostruisciPaccoRegolare
    RefreshDati
    DisegnoPaccoTuboRefresh
    Combo2.Visible = False
    ImgProfilo.Visible = Combo2.Visible And Combo2.ListIndex > 0
End Sub
Private Sub OptionPaccoSQSQ_Click()
    If Ricetta.TipoPacco = 1 Then Ricetta.NumeroFile = Ricetta.TubiFila(1)
    Ricetta.TipoPacco = Pacco.Quadro
    Ricetta.TipoTubo = Tubo.Quadro
    Call ControlloDatiPacco
    If FromOperator Then Ricetta.CostruisciPaccoRegolare
    RefreshDati
    DisegnoPaccoTuboRefresh
    Combo2.Visible = Param.GetBit("Par201_AbilitazioneProfili")
    ImgProfilo.Visible = Combo2.Visible And Combo2.ListIndex > 0
    If Combo2.ListIndex > 0 Then ImgProfilo.Picture = LoadPicture("..\Bitmap\" & NomeProfili(Combo2.ListIndex) & ".gif")
End Sub
' FINE - OPTION PACCO-TUBO

Private Sub ComModFile_Click()
    Dim i As Integer
    For i = 1 To MAX_ROWS
        ModFila(i) = Ricetta.TubiFila(i)
    Next
    FrameModFile.Visible = True
    FrameRecipeName.Visible = False
    OkCommand.Visible = False
    CancelCommand.Visible = False
    ModFileDisegno.Visible = True
End Sub

'------------ FrameModFile

Private Sub ModFileDisegno_Click()
    FrameModFile.Visible = False
    ModFileDisegno.Visible = False
    OkCommand.Visible = True
    CancelCommand.Visible = True
    If Param.GetBit("Par220_AttivaGestioneRicette") Then
       FrameRecipeName.Visible = True
       FrameRecipeName.ZOrder
    End If
End Sub

Private Sub ModFila_Click(Index As Integer)
    Dim i As Integer
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = ModFila(Index).Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        ModFila(Index).Text = TOUCHNumericPad.Dati
        For i = 1 To MAX_ROWS
            Ricetta.TubiFila(i) = ModFila(i)
        Next
        Ricetta.ControlloDatiPacco
        ' ricarica i dati con eventuali file modificate dal controllo
        For i = 1 To MAX_ROWS
            ModFila(i) = Ricetta.TubiFila(i)
        Next

        UpDownFilaBasePacco.value = Ricetta.TubiFila(1)
        UpDownFilaBasePacco.Refresh
        
        OggettoUpDownFile.value = Ricetta.NumeroFile
        OggettoUpDownFile.Refresh
        
        DisegnoPaccoTuboRefresh
        
    End If
End Sub

'***********************************************************
' FINE - Funzioni di risposta ai comandi dell'operatore ok e cancel
'***********************************************************
Private Sub DisplayLarghezza_CLICK()
Dim a
    a = Param.GetNumber("Par002_Tubo_LarghezzaMin")
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par002_Tubo_LarghezzaMin"), UM.mt, UM.inch, 3) ' * 1000
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par001_Tubo_LarghezzaMax"), UM.mt, UM.inch, 3) '* 1000
    TOUCHNumericPad.Dati = DisplayLarghezza.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.TuboLarghezza = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.inch, UM.mt, 6)
        Ricetta.ControlloDatiPacco
        RefreshDati
        AggiornaDaUpDownControl
    End If
End Sub

Private Sub DisplayAltezza_Click()
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par004_Tubo_AltezzaMin"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par003_Tubo_AltezzaMax"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.Dati = DisplayAltezza.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.TuboAltezza = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.inch, UM.mt, 6)
        Ricetta.ControlloDatiPacco
        RefreshDati
        AggiornaDaUpDownControl
    End If
End Sub

Private Sub DisplaySpessore_CLICK()
Dim a
     
    TOUCHNumericPad.Decimali = 4
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(Param.GetNumber("Par006_Tubo_SpessoreMin"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Param.GetNumber("Par005_Tubo_SpessoreMax"), UM.mt, UM.inch, 3)
    TOUCHNumericPad.Dati = DisplaySpessore.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.TuboSpessore = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.inch, UM.mt, 6)
        Ricetta.ControlloDatiPacco
        RefreshDati
    End If
End Sub

Private Sub DisplayLunghezza_CLICK()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = Fix(Conv_UM.Conversione(Param.GetNumber("Par008_Tubo_LunghezzaMin"), UM.mt, UM.ft))
    TOUCHNumericPad.ValoreMax = Fix(Conv_UM.Conversione(Param.GetNumber("Par007_Tubo_LunghezzaMax"), UM.mt, UM.ft))
    TOUCHNumericPad.Dati = DisplayLunghezza.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.TuboLunghezza = TOUCHNumericPad.Dati * 12 * 25.4 / 1000 + DisplayLungIN.caption * 25.4 / 1000
        Ricetta.ControlloDatiPacco
        Ricetta.CalcoloQuoteRegge
        RefreshDati
    End If
End Sub

Private Sub DisegnoPaccoTuboRefresh()
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
    If Ricetta.TipoTubo = Tubo.Quadro Then
        TuboModRicetta.Left = 1000
    Else
        TuboModRicetta.Left = 2200
        Ricetta.TuboAltezza = Ricetta.TuboLarghezza
    End If
    
    TuboModRicetta.aConfig = Param.GetNumber("Par101_MisureMetriche") + 2 + 4 + 0 'mm+tubo+spessore+label
    TuboModRicetta.aTube_Height = Conv_UM.Conversione(Ricetta.TuboAltezza, UM.mt, UM.inch, 2) * 100
    TuboModRicetta.aTube_Width = Conv_UM.Conversione(Ricetta.TuboLarghezza, UM.mt, UM.inch, 2) * 100
    TuboModRicetta.aTube_Tickness = Conv_UM.Conversione(Ricetta.TuboSpessore, UM.mt, UM.inch, 4) * 1000
    TuboModRicetta.aTipoTubo = Ricetta.TipoTubo
    
    'PACCO
    PaccoModRicetta.aConfig = Param.GetNumber("Par101_MisureMetriche") + 0 + 4 + 8 + Abs(Ricetta.Profilo > 0) * 16 'mm+pacco+spessore+label
    PaccoModRicetta.aTube_Height = Conv_UM.Conversione(Ricetta.TuboAltezza, UM.mt, UM.inch, 2) * 100
    PaccoModRicetta.aTube_Width = Conv_UM.Conversione(Ricetta.TuboLarghezza, UM.mt, UM.inch, 2) * 100
    PaccoModRicetta.aTube_Tickness = Conv_UM.Conversione(Ricetta.TuboSpessore, UM.mt, UM.inch, 4) * 1000
    PaccoModRicetta.aTipoTubo = Ricetta.TipoTubo
    PaccoModRicetta.aTipoPacco = Ricetta.TipoPacco
    PaccoModRicetta.aTipoProfilo = Ricetta.Profilo
    
    '========================================================================
    lblModOrdine.caption = Ricetta.NumeroTubiPacco
    Ricetta.Pieces = Ricetta.NumeroTubiPacco
'    LblPesoModOrdine.caption = Conv_UM.Conversione(Ricetta.TuboPesoTeorico * Ricetta.NumeroTubiPacco, UM.kg, UM.lB)
    LblPesoModOrdine.caption = Round(Conv_UM.Conversione(Ricetta.TuboLunghezza, UM.mt, UM.ft, 0) * Ricetta.WeightPerFeet * Ricetta.NumeroTubiPacco, 0)
    Ricetta.Weight = LblPesoModOrdine
    Label14 = Ricetta.WeightPerFeet
    '========================================================================
    
    PaccoModRicetta.aCounted = 0
    
    Dim i As Integer
    For i = 1 To MAX_ROWS
        PaccoModRicetta.TubiFila(i) = Ricetta.TubiFila(i)
    Next
    
    TuboModRicetta.Refresh
    PaccoModRicetta.Refresh
    
End Sub

Public Sub AggiornaDaUpDownControl()
   Static LocTubifila, LocNumFile
    
      ' assegna il numero di tubi in fila base
    Ricetta.TubiFila(1) = UpDownFilaBasePacco.value
    Ricetta.NumeroFile = OggettoUpDownFile.value
    ' controlla i limiti del pacco
    Call ControlloDatiPacco
    
    'aggiornamento pacco regolare
    If SSTab1.Tab <> 1 Then Ricetta.CostruisciPaccoRegolare
    DisegnoPaccoTuboRefresh
    
    'aggiornamento controlli updown
    
    ControlloVelTR.LimMin = 20
    ControlloVelTR.LimMax = 80
    Ricetta.VelTR = ControlloVelTR.value
    ControlloVelTR.Refresh
    
    ControlloModMagneti.LimMin = 20
    ControlloModMagneti.LimMax = 100
    Ricetta.VelMPS = ControlloModMagneti.value
    ControlloModMagneti.Refresh
    
'    ControlloOverrideVR1.Step = 10
'    ControlloOverrideVR1.LimMin = 20
'    ControlloOverrideVR1.LimMax = 100
'    Ricetta.VelVR1 = ControlloOverrideVR1.value
'    ControlloOverrideVR1.Refresh
'    ControlloOverrideVR2.Step = 10
'    ControlloOverrideVR2.LimMin = 20
'    ControlloOverrideVR2.LimMax = 100
'    Ricetta.VelVR2 = ControlloOverrideVR2.value
'    ControlloOverrideVR2.Refresh
'    ControlloMonobeam1.Step = 10
'    ControlloMonobeam1.LimMin = 20
'    ControlloMonobeam1.LimMax = 100
'    Ricetta.VelMB1 = ControlloMonobeam1.value
'    ControlloMonobeam1.Refresh
'    ControlloMonobeam2.Step = 10
'    ControlloMonobeam2.LimMin = 20
'    ControlloMonobeam2.LimMax = 100
'    Ricetta.VelMB2 = ControlloMonobeam2.value
'    ControlloMonobeam2.Refresh
'    ControlloMonobeam3.Step = 10
'    ControlloMonobeam3.LimMin = 20
'    ControlloMonobeam3.LimMax = 100
'    Ricetta.VelMB3 = ControlloMonobeam3.value
'    ControlloMonobeam3.Refresh
    
    Ricetta.NumeroRegge = ControlloUpDownRegge.value
    ControlloUpDownRegge.Refresh
    Ricetta.CalcoloQuoteRegge
    RefreshRegge
End Sub

Private Sub Selettore_Click(Index As Integer)
    If Index = 2 Then Ricetta.Regg1 = Not Ricetta.Regg1
    If Index = 3 Then Ricetta.Regg2 = Not Ricetta.Regg2
    If Ricetta.Regg1 = False And Ricetta.Regg2 = False Then Ricetta.Regg1 = True
    Selettore(2).Picture = ImageList1.ListImages(1 + Abs(CInt(Ricetta.Regg1))).Picture
    Selettore(3).Picture = ImageList1.ListImages(1 + Abs(CInt(Ricetta.Regg2))).Picture
End Sub

Private Sub SelettoreSagome_Click()
   ModOrder.TicketUnit = Not ModOrder.TicketUnit
   SelettoreSagome.Picture = ImageList1.ListImages(Abs(ModOrder.TicketUnit) + 1).Picture
   Cartellino.UnitMisura = IIf(Cartellino.UnitMisura = metrica, TUnitMis.inch, TUnitMis.metrica)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    DisegnoLunghezzaRefresh
End Sub

Private Sub Text1_Click(Index As Integer)
    If Index = 2 Then
        DialogSerials.Show vbModal
        If Trim(DialogSerials.SerialNumber) <> "" And Trim(DialogSerials.RecipeName) <> "" Then
'           If OrderModifyForm.NewOrModify Then Ricetta.IDRicetta = Trim(DialogSerials.RecipeName)
           Cartellino.CampoManuale(2) = Trim(DialogSerials.GradeNumber)
           Text1(Index).Text = Cartellino.CampoManuale(2)
        End If
        Unload DialogSerials
    Else
        TOUCHKeyBoard.TextModifica.PasswordChar = ""
        TOUCHKeyBoard.Dati = Text1(Index).Text
        TOUCHKeyBoard.Show vbModal
        If TOUCHKeyBoard.DatiConfermati Then
            Text1(Index).Text = TOUCHKeyBoard.Dati
            If Index > 10 Then
               Cartellino.CampoAuto(Index - 10) = TOUCHKeyBoard.Dati
            Else
               Cartellino.CampoManuale(Index) = TOUCHKeyBoard.Dati
            End If
        End If
        If Index = 18 Then
           WeightToPrint = TOUCHKeyBoard.Dati
        End If
        TOUCHKeyBoard.TextModifica.PasswordChar = ""
    End If
End Sub

Private Sub TextReggia_CLICK(Index As Integer)
    TOUCHNumericPad.Decimali = 2
    TOUCHNumericPad.ValoreMin = Conv_UM.Conversione(0.001, UM.mt, UM.inch)
    TOUCHNumericPad.ValoreMax = Conv_UM.Conversione(Ricetta.TuboLunghezza, UM.mt, UM.inch)
    TOUCHNumericPad.Dati = TextReggia(Index).caption
    TOUCHNumericPad.Show vbModal
     
    If TOUCHNumericPad.DatiConfermati Then
        Ricetta.QuotaReggia(Index) = Conv_UM.Conversione(TOUCHNumericPad.Dati, UM.inch, UM.mt, 6)    ' NuovoValore
        TextReggia(1).caption = TOUCHNumericPad.Dati
    End If
    
    If Index = 1 And Ricetta.TipoCalcRegge = 0 Then
        ControlloUpDownRegge.Cliccato = True
        AggiornaDaUpDownControl
'        Ricetta.CalcoloQuoteRegge
    End If
End Sub

Private Function CheckReggia(Index As Integer, valore As Double) As Boolean
    CheckReggia = True
    ' se è l'ultima reggia la devo limitare alla lunghezza del tubo
    If Ricetta.NumeroRegge = Index Then
        If valore > Ricetta.TuboLunghezza Then
            CheckReggia = False
        End If
    End If
    ' per le regge intermedie si controlla che la posizione non superi la reggia successuva
    If Ricetta.NumeroRegge > Index Then
        If valore > Ricetta.QuotaReggia(Index + 1) Then
            CheckReggia = False
        End If
    End If
    
    ' la prima reggia non deve essere in posizione inferiore a 0
    If Index = 1 Then
        If valore < 0 Then
            CheckReggia = False
        End If
    End If
    
    ' le reggie successive alla prima devono essere superiori alla reggia precedente
    If Index > 1 Then
        If valore < Ricetta.QuotaReggia(Index - 1) Then
            CheckReggia = False
        End If
    End If
End Function

Private Sub Timer1_Timer()
Static Lamp As Boolean

  Lamp = Not Lamp
  If Lamp Then
     Image1.Visible = True
  Else
     Image1.Visible = False
  End If
End Sub

Sub ScritteMultilingua()
    FramePacco.caption = Param.Text("Pacco")
    FrameTubo.caption = Param.Text("Tubo")
    FrameRicette.caption = Param.Text("ORDER")
    FrameRecipeName.caption = Param.Text("Ricette")
    CancelCommand.caption = Param.Text("Annulla")
    ComModFile.caption = Param.Text("SPECIALE")
    SSTab1.TabCaption(0) = Param.Text("Bundle")
    SSTab1.TabCaption(1) = Param.Text("Strap")
    SSTab1.TabCaption(2) = Param.Text("Generale")
    SSTab1.TabCaption(3) = Param.Text("MPS")
    SSTab1.TabCaption(4) = Param.Text("Ticket")
    SSTab2.TabCaption(1) = Param.Text("Automatici")
    SSTab2.TabCaption(0) = Param.Text("Manuali")
    CmdPercorso.caption = Param.Text("Percorso")
    Command1(0).caption = Param.Text("Caricatore")
    Command1(1).caption = Param.Text("Verniciatrice")
    Label8.caption = Param.Text("Vel. magneti (%)")
    ModFileDisegno.caption = Param.Text("Disegno")
    Label1.caption = Param.Text("Dopo")
    Label2.caption = Param.Text("Pacchi")
    Combo1.Clear
    Combo1.AddItem Param.Text("Non stop")
    Combo1.AddItem Param.Text("Arresto")
    Combo1.AddItem Param.Text("Nuovo ordine")
    Label6.caption = Param.Text("Entrata")
    Label3.caption = Param.Text("Ntubi")
    Label4.caption = Param.Text("Peso2")
    Label5.caption = Param.Text("VelTRSalita")
'    Label9.caption = Param.Text("Velocità via rulli (s)") & " 1 (%)"
'    Label10.caption = Param.Text("Velocità via rulli (s)") & " 2 (%)"
'    Label11.caption = Param.Text("VelMonobeam") & " 1 (%)"
'    Label12.caption = Param.Text("VelMonobeam") & " 2 (%)"
'    Label13.caption = Param.Text("VelMonobeam") & " 3 (%)"
    Label16.caption = Param.Text("ReggModo")
    If Param.GetBit("Par101_MisureMetriche") Then
       Label20.caption = Param.Text("Reggeoffset") & " [mm]"
    Else
       Label20.caption = Param.Text("Reggeoffset") & " [inch]"
    End If
    Label21 = Param.Text("Nticket")
End Sub

Sub ControlloDatiPacco()
    Dim LargMaxPacco As Double
    Dim LargMinPacco As Double
    Dim AltezzaMinPacco As Double
    Dim AltezzaMaxPacco As Double

    '=============== controllo dati pacco =================================
    
    'calcola il numero di tubi in fila base e il numero di file
    Ricetta.LarghezzaMaxPacco = Param.GetNumber("Par023_Pacco_LarghezzaMaxBaseSQ") / 1000
    Ricetta.LarghezzaMinPacco = Param.GetNumber("Par024_Pacco_LarghezzaMinBaseSQ") / 1000
    Ricetta.AltezzaMaxPacco = Param.GetNumber("Par027_Pacco_AltezzaMax") / 1000
    Ricetta.AltezzaMinPacco = Param.GetNumber("Par021_Pacco_AltezzaMin") / 1000
    Ricetta.LarghezzaMaxPaccoEsa = Param.GetNumber("Par021_Pacco_LarghezzaMaxBaseHex") / 1000
    Ricetta.LarghezzaMinPaccoEsa = Param.GetNumber("Par022_Pacco_LarghezzaMinBaseHex") / 1000
    'ricalcolo pacco
    Ricetta.CalcoloPaccoLarghezza
    Ricetta.CalcoloPaccoAltezza
    
    '********** controllo larghezza *************
    
    ' MIN quadro
    If (Ricetta.NumMinTubiFila > Ricetta.TubiFila(1) And Ricetta.TipoPacco = Pacco.Quadro) Then
          UpDownFilaBasePacco.value = Ricetta.NumMinTubiFila
          Ricetta.TubiFila(1) = Ricetta.NumMinTubiFila
          UpDownFilaBasePacco.Refresh
    End If
    'MIN esagono
     If (Ricetta.NumMinTubiFilaEsa > Ricetta.TubiFila(1) And Ricetta.TipoPacco = Pacco.Esagono) Then
          UpDownFilaBasePacco.value = Ricetta.NumMinTubiFilaEsa
          Ricetta.TubiFila(1) = Ricetta.NumMinTubiFilaEsa
          UpDownFilaBasePacco.Refresh
    End If
    'MAX quadro
    If (Ricetta.NumMaxTubiFila < Ricetta.TubiFila(1) And Ricetta.TipoPacco = Pacco.Quadro) Then
          UpDownFilaBasePacco.value = Ricetta.NumMaxTubiFila
          Ricetta.TubiFila(1) = Ricetta.NumMaxTubiFila
          UpDownFilaBasePacco.Refresh
    End If
     'MAX esagono
    If (Ricetta.NumMaxTubiFilaEsa < Ricetta.TubiFila(1) And Ricetta.TipoPacco = Pacco.Esagono) Then
          UpDownFilaBasePacco.value = Ricetta.NumMaxTubiFilaEsa
          Ricetta.TubiFila(1) = Ricetta.NumMaxTubiFilaEsa
          UpDownFilaBasePacco.Refresh
    End If
    'controllo altezza
    If Ricetta.NumMinFile > Ricetta.NumeroFile And Ricetta.TipoPacco = Pacco.Quadro Then
          OggettoUpDownFile.value = Ricetta.NumMinFile
          Ricetta.NumeroFile = Ricetta.NumMinFile
          OggettoUpDownFile.Refresh
    End If
    If Ricetta.NumMaxFile < Ricetta.NumeroFile And Ricetta.TipoPacco = Pacco.Quadro Then
          OggettoUpDownFile.value = Ricetta.NumMaxFile
          Ricetta.NumeroFile = Ricetta.NumMaxFile
          OggettoUpDownFile.Refresh
    End If
    '========================================================================
    
End Sub

