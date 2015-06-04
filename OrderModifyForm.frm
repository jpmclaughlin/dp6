VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form OrderModifyForm 
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
   Begin VB.CommandButton Command3 
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
      TabIndex        =   81
      Top             =   10560
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.CommandButton CloseCommand 
      BackColor       =   &H0000FF00&
      Caption         =   "Chiudi"
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
      Height          =   795
      Left            =   12630
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   10590
      Width           =   2300
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7320
      Top             =   10800
   End
   Begin VB.CommandButton OkCommand 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10590
      Width           =   2295
   End
   Begin VB.CommandButton CancelCommand 
      BackColor       =   &H000000FF&
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
      Height          =   795
      Left            =   12630
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10590
      Width           =   2300
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
            Picture         =   "OrderModifyForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrderModifyForm.frx":0603
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   1905
      Left            =   -60
      TabIndex        =   75
      Top             =   -180
      Width           =   15420
      _ExtentX        =   27199
      _ExtentY        =   3360
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "OrderModifyForm.frx":0C15
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "Operazione"
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
         Height          =   1275
         Left            =   120
         TabIndex        =   76
         Top             =   150
         Width           =   15075
         Begin VB.CommandButton CmdNuovaRicetta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nuova"
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
            Left            =   3930
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   450
            Width           =   2325
         End
         Begin VB.CommandButton CmdModificaRicetta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modifica"
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
            Left            =   6630
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   450
            Width           =   2325
         End
         Begin VB.CommandButton CmdCancellaRicetta 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancella"
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
            Left            =   9330
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   450
            Width           =   2355
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10485
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   18494
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1235
      BackColor       =   8421504
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
      TabCaption(0)   =   "ORDER"
      TabPicture(0)   =   "OrderModifyForm.frx":0C31
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameOrdine"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameOrderEnd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameSetup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ticket"
      TabPicture(1)   =   "OrderModifyForm.frx":0C4D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label24"
      Tab(1).Control(1)=   "Label01(1)"
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(3)=   "SelettoreSagome"
      Tab(1).Control(4)=   "SSTab2"
      Tab(1).Control(5)=   "OptionLingua(2)"
      Tab(1).Control(6)=   "OptionLingua(4)"
      Tab(1).Control(7)=   "OptionLingua(3)"
      Tab(1).Control(8)=   "OptionLingua(1)"
      Tab(1).Control(9)=   "OptionLingua(0)"
      Tab(1).Control(10)=   "OptionLingua(5)"
      Tab(1).ControlCount=   11
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
         Left            =   -62340
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   5400
         Visible         =   0   'False
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
         Left            =   -62370
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1410
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
         Left            =   -62370
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2220
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
         Left            =   -62370
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3000
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
         Left            =   -62370
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3780
         Visible         =   0   'False
         Width           =   2355
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
         Left            =   -62340
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4590
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Frame FrameSetup 
         Caption         =   "SETUP"
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
         Height          =   8280
         Left            =   4320
         TabIndex        =   5
         Top             =   2070
         Width           =   10830
         Begin VB.Frame Frame2 
            Height          =   855
            Left            =   6690
            TabIndex        =   85
            Top             =   90
            Width           =   4125
            Begin VB.Label Label8 
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
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   2490
               TabIndex        =   87
               Top             =   210
               Width           =   1470
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Storage dest."
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
               Left            =   90
               TabIndex        =   86
               Top             =   270
               Width           =   2280
            End
         End
         Begin dp6.ControlloRegge ControlloRegge 
            Height          =   2070
            Left            =   1260
            TabIndex        =   8
            Top             =   5400
            Width           =   8805
            _ExtentX        =   15531
            _ExtentY        =   1905
         End
         Begin dp6.ControlloPacco TuboModOrdine 
            Height          =   3945
            Left            =   6330
            TabIndex        =   7
            Top             =   1230
            Width           =   2850
            _ExtentX        =   6932
            _ExtentY        =   8281
         End
         Begin dp6.ControlloPacco PaccoModOrdine 
            Height          =   4485
            Left            =   1740
            TabIndex        =   6
            Top             =   690
            Width           =   3975
            _ExtentX        =   7752
            _ExtentY        =   8573
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
            Left            =   3540
            TabIndex        =   70
            Top             =   7650
            Width           =   1725
         End
         Begin VB.Label Label7 
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
            Height          =   555
            Left            =   1290
            TabIndex        =   69
            Top             =   7650
            Width           =   4065
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
            Height          =   675
            Left            =   7440
            TabIndex        =   67
            Top             =   7650
            Width           =   2235
         End
         Begin VB.Image ImgProfilo 
            Height          =   2205
            Left            =   8040
            Top             =   2190
            Width           =   2655
         End
         Begin VB.Label Label6 
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
            Height          =   570
            Left            =   6000
            TabIndex        =   68
            Top             =   7650
            Width           =   3735
         End
      End
      Begin VB.Frame FrameOrderEnd 
         Caption         =   "Recipe Select"
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
         Height          =   8250
         Left            =   90
         TabIndex        =   3
         Top             =   2070
         Width           =   4140
         Begin VB.Frame framefilter 
            Height          =   855
            Left            =   180
            TabIndex        =   72
            Top             =   7350
            Width           =   3735
            Begin VB.CommandButton cmdFiltro 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   2580
               TabIndex        =   74
               Top             =   210
               Width           =   1035
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Filtro"
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
               Height          =   495
               Left            =   690
               TabIndex        =   73
               Top             =   270
               Width           =   2385
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   60
               Picture         =   "OrderModifyForm.frx":0C69
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "UP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   450
            Width           =   3135
         End
         Begin VB.CommandButton Command1 
            Caption         =   "DOWN"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   510
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   6780
            Width           =   3135
         End
         Begin VB.ListBox ListaRicette 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5820
            Left            =   180
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   960
            Width           =   3735
         End
         Begin VB.Label LabelDopoPacchi 
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
            Left            =   75
            TabIndex        =   4
            Top             =   5145
            Width           =   3810
         End
      End
      Begin VB.Frame FrameOrdine 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1275
         Left            =   60
         TabIndex        =   12
         Top             =   780
         Width           =   15135
         Begin VB.TextBox Text1 
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
            Index           =   0
            Left            =   13080
            TabIndex        =   14
            Text            =   "100"
            Top             =   255
            Width           =   1575
         End
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
            Left            =   8280
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   300
            Width           =   3075
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
            Height          =   720
            Left            =   180
            TabIndex        =   17
            Top             =   240
            Width           =   7695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Dopo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   11520
            TabIndex        =   16
            Top             =   195
            Width           =   1470
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pacchi"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Left            =   11520
            TabIndex        =   15
            Top             =   555
            Width           =   1470
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   9315
         Left            =   -74850
         TabIndex        =   19
         Top             =   1080
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   16431
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   1058
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
         TabPicture(0)   =   "OrderModifyForm.frx":0F73
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "campo(3)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "campo(4)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "campo(5)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "campo(6)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "campo(7)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "campo(8)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "campo(9)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "campo(10)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "campo(2)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "campo(1)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Line1(9)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Line1(8)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Line1(7)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Line1(6)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Line1(5)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Line1(4)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Line1(3)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Line1(2)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Line1(1)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Line1(0)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "campo(0)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "LblCartellini"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "Label5"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Text1(3)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "Text1(4)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "Text1(5)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "Text1(6)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "Text1(7)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Text1(8)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "Text1(9)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Text1(10)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "Text1(2)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Text1(1)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).ControlCount=   33
         TabCaption(1)   =   "Automatici"
         TabPicture(1)   =   "OrderModifyForm.frx":0F8F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "campo(11)"
         Tab(1).Control(1)=   "campo(12)"
         Tab(1).Control(2)=   "campo(20)"
         Tab(1).Control(3)=   "campo(19)"
         Tab(1).Control(4)=   "campo(18)"
         Tab(1).Control(5)=   "campo(17)"
         Tab(1).Control(6)=   "campo(16)"
         Tab(1).Control(7)=   "campo(15)"
         Tab(1).Control(8)=   "campo(14)"
         Tab(1).Control(9)=   "campo(13)"
         Tab(1).Control(10)=   "Line1(10)"
         Tab(1).Control(11)=   "Line1(11)"
         Tab(1).Control(12)=   "Line1(12)"
         Tab(1).Control(13)=   "Line1(13)"
         Tab(1).Control(14)=   "Line1(14)"
         Tab(1).Control(15)=   "Line1(15)"
         Tab(1).Control(16)=   "Line1(16)"
         Tab(1).Control(17)=   "Line1(17)"
         Tab(1).Control(18)=   "Line1(18)"
         Tab(1).Control(19)=   "Line1(19)"
         Tab(1).Control(20)=   "Text1(11)"
         Tab(1).Control(21)=   "Text1(12)"
         Tab(1).Control(22)=   "Text1(20)"
         Tab(1).Control(23)=   "Text1(19)"
         Tab(1).Control(24)=   "Text1(18)"
         Tab(1).Control(25)=   "Text1(17)"
         Tab(1).Control(26)=   "Text1(16)"
         Tab(1).Control(27)=   "Text1(15)"
         Tab(1).Control(28)=   "Text1(14)"
         Tab(1).Control(29)=   "Text1(13)"
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
            Index           =   13
            Left            =   -69720
            TabIndex        =   39
            Top             =   2040
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
            TabIndex        =   38
            Top             =   2730
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
            TabIndex        =   37
            Top             =   3420
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
            TabIndex        =   36
            Top             =   4110
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
            TabIndex        =   35
            Top             =   4800
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
            TabIndex        =   34
            Top             =   5490
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
            Left            =   -69720
            TabIndex        =   33
            Top             =   6180
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
            TabIndex        =   32
            Top             =   6870
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
            Left            =   -69720
            TabIndex        =   31
            Top             =   1350
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
            TabIndex        =   30
            Top             =   660
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
            TabIndex        =   29
            Top             =   1350
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
            Index           =   2
            Left            =   5280
            TabIndex        =   28
            Top             =   2700
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
            Index           =   10
            Left            =   5280
            TabIndex        =   27
            Top             =   4110
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
            TabIndex        =   26
            Top             =   7560
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
            TabIndex        =   25
            Top             =   6870
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
            TabIndex        =   24
            Top             =   6180
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
            TabIndex        =   23
            Top             =   5490
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
            Index           =   5
            Left            =   5280
            TabIndex        =   22
            Top             =   4800
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
            Index           =   4
            Left            =   5130
            TabIndex        =   21
            Top             =   7560
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
            Index           =   3
            Left            =   5280
            TabIndex        =   20
            Top             =   3420
            Width           =   6375
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tickets"
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
            Height          =   645
            Left            =   5610
            TabIndex        =   84
            Top             =   7800
            Width           =   2205
         End
         Begin VB.Label LblCartellini 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   585
            Left            =   5610
            TabIndex        =   83
            Top             =   8430
            Width           =   2205
         End
         Begin VB.Label campo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Enter Customer Name, or Cut for Cut-to-Lenght, Or Leave blank for Stock"
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
            Height          =   645
            Index           =   0
            Left            =   5280
            TabIndex        =   82
            Top             =   2010
            Width           =   6375
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   0
            X1              =   360
            X2              =   5760
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   1
            X1              =   360
            X2              =   5760
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   420
            X2              =   5820
            Y1              =   6090
            Y2              =   6090
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   3
            X1              =   420
            X2              =   5820
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   4
            X1              =   420
            X2              =   5820
            Y1              =   4740
            Y2              =   4740
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   420
            X2              =   5820
            Y1              =   4050
            Y2              =   4050
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   6
            X1              =   360
            X2              =   5760
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   7
            X1              =   360
            X2              =   5760
            Y1              =   2670
            Y2              =   2670
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   8
            X1              =   420
            X2              =   5820
            Y1              =   6810
            Y2              =   6810
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   9
            X1              =   420
            X2              =   5820
            Y1              =   7500
            Y2              =   7500
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   19
            X1              =   -74670
            X2              =   -69270
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   18
            X1              =   -74580
            X2              =   -69180
            Y1              =   7500
            Y2              =   7500
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   17
            X1              =   -74580
            X2              =   -69180
            Y1              =   6840
            Y2              =   6840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   16
            X1              =   -74640
            X2              =   -69240
            Y1              =   2670
            Y2              =   2670
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   15
            X1              =   -74640
            X2              =   -69240
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   14
            X1              =   -74580
            X2              =   -69180
            Y1              =   4080
            Y2              =   4080
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   13
            X1              =   -74580
            X2              =   -69180
            Y1              =   4740
            Y2              =   4740
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   12
            X1              =   -74580
            X2              =   -69180
            Y1              =   5460
            Y2              =   5460
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   11
            X1              =   -74580
            X2              =   -69180
            Y1              =   6120
            Y2              =   6120
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00E0E0E0&
            Index           =   10
            X1              =   -74640
            X2              =   -69240
            Y1              =   1980
            Y2              =   1980
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
            Height          =   495
            Index           =   13
            Left            =   -74760
            TabIndex        =   59
            Top             =   2100
            Width           =   4755
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
            Height          =   495
            Index           =   14
            Left            =   -74760
            TabIndex        =   58
            Top             =   2790
            Width           =   4755
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
            Height          =   495
            Index           =   15
            Left            =   -74760
            TabIndex        =   57
            Top             =   3510
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
            Height          =   495
            Index           =   16
            Left            =   -74760
            TabIndex        =   56
            Top             =   4170
            Width           =   4755
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
            Height          =   495
            Index           =   17
            Left            =   -74760
            TabIndex        =   55
            Top             =   4860
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
            Height          =   495
            Index           =   18
            Left            =   -74760
            TabIndex        =   54
            Top             =   5550
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
            Height          =   495
            Index           =   19
            Left            =   -74760
            TabIndex        =   53
            Top             =   6210
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
            Index           =   20
            Left            =   -74760
            TabIndex        =   52
            Top             =   6930
            Width           =   4755
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
            Height          =   495
            Index           =   12
            Left            =   -74760
            TabIndex        =   51
            Top             =   1410
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
            Height          =   495
            Index           =   11
            Left            =   -74760
            TabIndex        =   50
            Top             =   750
            Width           =   4785
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
            Left            =   240
            TabIndex        =   49
            Top             =   1410
            Width           =   4695
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
            Left            =   240
            TabIndex        =   48
            Top             =   2730
            Width           =   4695
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
            Left            =   240
            TabIndex        =   47
            Top             =   4200
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
            Left            =   240
            TabIndex        =   46
            Top             =   7560
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
            Index           =   8
            Left            =   240
            TabIndex        =   45
            Top             =   6900
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
            Left            =   240
            TabIndex        =   44
            Top             =   6210
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
            Left            =   240
            TabIndex        =   43
            Top             =   5520
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
            Left            =   240
            TabIndex        =   42
            Top             =   4860
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
            Left            =   330
            TabIndex        =   41
            Top             =   7740
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
            Index           =   3
            Left            =   240
            TabIndex        =   40
            Top             =   3450
            Width           =   4755
         End
      End
      Begin VB.Image SelettoreSagome 
         Height          =   1275
         Left            =   -61740
         Top             =   7410
         Visible         =   0   'False
         Width           =   1275
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
         Height          =   330
         Left            =   -62250
         TabIndex        =   66
         Top             =   6390
         Visible         =   0   'False
         Width           =   2160
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
         Left            =   -62370
         TabIndex        =   65
         Top             =   7080
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   -62250
         TabIndex        =   71
         Top             =   6690
         Visible         =   0   'False
         Width           =   2175
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
' Costanti e variabili per gestione input dati
'*************************************************
' Ok è public perchè può essere esaminata per verificare
' se i dati sono stati confermati
Public OK As Boolean
'' flag protezione modifica  dati
'Private FromOperator As Boolean
'**************************************************
' Fine costanti e variabili per gestione input dati
'**************************************************

Private LinguaSelezionata As Byte

Public oneShot As Boolean
Public PulsantePremuto As Boolean
Public NewOrModify As Boolean

Private Sub CloseCommand_Click()
    OK = False
    oneShot = False
    Me.Hide
End Sub

Private Sub CmdCancellaRicetta_Click()
Dim Segno As Integer
Dim DaCancellare As Boolean

    DaCancellare = False
    
    With OrdersForm.AdoOrdini
           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
           .CommandType = adCmdTable
           .RecordSource = "Recipes"
           .Refresh
           .Recordset.MoveFirst
             
           While .Recordset.EOF = False
               If .Recordset.Fields("ID") = ModRecipe.IDRicetta Then
                  OrderDeleteForm.OrderLabel = ModRecipe.IDRicetta
                  OrderDeleteForm.RecipeOrOrder = True
                  OrderDeleteForm.Show vbModal
                  If OrderDeleteForm.OK Then
                     .RecordSource = "Orders"
                     .Refresh
                     .Recordset.MoveFirst
                     DaCancellare = True
                      While .Recordset.EOF = False
                          If .Recordset.Fields("IDRicetta") = ModRecipe.IDRicetta Then
                              If .Recordset.Fields("Visualizzato") Then DaCancellare = False
                              If frmKernel.CodOrdineCorrente.CodEntrata = .Recordset.Fields("ID") Then DaCancellare = False
                              If frmKernel.CodOrdineCorrente.CodMPS = .Recordset.Fields("ID") Then DaCancellare = False
                              If frmKernel.CodOrdineCorrente.CodPacco = .Recordset.Fields("ID") Then DaCancellare = False
                              If frmKernel.CodOrdineCorrente.CodRegge = .Recordset.Fields("ID") Then DaCancellare = False
                              If frmKernel.CodOrdineCorrente.CodStoccaggio = .Recordset.Fields("ID") Then DaCancellare = False
                              If frmKernel.CodOrdineCorrente.CodWB = .Recordset.Fields("ID") Then DaCancellare = False
                          End If
                          .Recordset.MoveNext
                      Wend
                     .RecordSource = "Recipes"
                     .Refresh
                     .Recordset.MoveFirst
                     .Recordset.Find "ID='" & ModRecipe.IDRicetta & "'"
                     If DaCancellare Then
                        .Recordset.Delete
                     End If
                  Else
                      Set .Recordset.ActiveConnection = Nothing
                      Exit Sub
                  End If
               End If
               .Recordset.MoveNext
           Wend
           Set .Recordset.ActiveConnection = Nothing
           RefreshListaRicette
        End With
End Sub

Private Sub cmdFiltro_Click()
   DoEvents
   frmFilter.Show vbModal
   DoEvents
   If frmFilter.Filtrostr = "" Then Exit Sub
   RefreshListaRicette frmFilter.Filtrostr
End Sub

Private Sub CmdModificaRicetta_Click()
     RecipeModifyForm.SSTab1.Tab = 0
     RecipeModifyForm.Update
     OrdersForm.PaginaRicette
     If RecipeModifyForm.OK Then
        OrdersForm.SaveRecipeData Ricetta
     End If
     oneShot = False
End Sub

Private Sub CmdNuovaRicetta_Click()
    Ricetta.IDRicetta = ModRecipe.IDRicetta
    OrdersForm.LoadRecipeData Ricetta
    Ricetta.IDRicetta = "????????????"
    NewOrModify = True
    RecipeModifyForm.SSTab1.Tab = 0
    RecipeModifyForm.Update
    OrdersForm.PaginaRicette
    If RecipeModifyForm.OK Then
       OrdersForm.SaveRecipeData Ricetta
    End If
    oneShot = False
End Sub

Private Sub Combo1_Click()
 Select Case Combo1.ListIndex
      Case 0
               ModOrder.ModoCambioOrdine = 0
               Label1.Visible = False
               Label2.Visible = False
               Text1(0).Visible = False
      Case 1
               ModOrder.ModoCambioOrdine = 1
               Label1.Visible = True
               Label2.Visible = True
               Text1(0).Visible = True
      Case 2
               ModOrder.ModoCambioOrdine = 2
               Label1.Visible = True
               Label2.Visible = True
               Text1(0).Visible = True
     End Select
    ' PulsantePremuto = True
End Sub

Private Sub Command1_Click()
      If ListaRicette.ListIndex < ListaRicette.ListCount - 1 Then
         ListaRicette.ListIndex = ListaRicette.ListIndex + 1
      End If
End Sub

Private Sub Command2_Click()
      If ListaRicette.ListIndex > 0 Then
         ListaRicette.ListIndex = ListaRicette.ListIndex - 1
      End If
End Sub

Private Sub Command3_Click()
   frmStampa.RefreshVar
   frmStampa.Command1.Visible = True
   frmStampa.Show vbModal
   frmStampa.Command1.Visible = False
End Sub

'===============================================================================
'                                    INIZIO FUNZIONI PAGINA
'===============================================================================
Private Sub Form_Load()
    ScritteMultilingua
    PulsantePremuto = False
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    If oneShot Then Exit Sub
    OK = False
    NewOrModify = False
  '  Picture1.Visible = False
    SelettoreSagome.Picture = ImageList1.ListImages(1 + Abs(CInt(ModOrder.TicketUnit))).Picture
    LblCartellini.caption = Param.GetNumber("Par221_NumeroCartellini")
    If OrdersForm.CmdRicetteEnable Then
      FrameOrdine.Visible = False
      CancelCommand.Visible = False
      OkCommand.Visible = False
      CloseCommand.Visible = True
      Frame1.Visible = True
      SSTab3.TabCaption(0) = ""
      SSTab3.Visible = True
     ' FrameRicette.Visible = True
    Else
      FrameOrdine.Visible = True
      CancelCommand.Visible = True
      OkCommand.Visible = True
      CloseCommand.Visible = False
      Frame1.Visible = False
      SSTab3.TabCaption(0) = ""
      SSTab3.Visible = False
     ' FrameRicette.Visible = False
    End If
    Combo1.Text = Combo1.List(ModOrder.ModoCambioOrdine)
    Text1(0).Text = ModOrder.PresetPacchi
    DisplayDescrizione.caption = ModOrder.Descrizione
    PaccoModOrdine.ColoreSfondo = &H8000000F
    TuboModOrdine.ColoreSfondo = &H8000000F
    Text1(0).Visible = False
     If ModOrder.ModoCambioOrdine > 0 Then
       Text1(0).Visible = True
    End If
    'ModRecipe
    'Disegno pacco-tubo
    Call DisegnoPaccoTuboRefresh
    
    'Disegno regge
    Call DisegnoReggeRefresh
    If PulsantePremuto Then Call RefreshListaRicette
   
    PulsantePremuto = False
    
    ' stampa
    SSTab2.TabEnabled(1) = False
    If Param.GetBit("Par204_AttivaPesa") Then
       SSTab1.TabEnabled(1) = Param.GetBit("Par208_PresenzaStampante") And PrinterInstall
       SSTab2.Tab = 0
       '========================================================================
       ' aggiornamento campi manuali del cartellino con quelli contenuti nella ricetta
       '========================================================================
       ModOrder.CampoManuale(4) = ModRecipe.IDRicetta
       ModOrder.CampoManuale(2) = ModRecipe.Grade
       ModOrder.CampoManuale(5) = ModRecipe.Pieces
       ModOrder.CampoManuale(8) = ModRecipe.Itemcode
       ModOrder.CampoManuale(7) = ModRecipe.Weight
       ' aggiorna il cartellino
       Cartellino.UnitMisura = IIf(ModOrder.TicketUnit, TUnitMis.inch, TUnitMis.metrica)
       For i = 1 To 20
           If i < 11 Then
              Cartellino.CampoManuale(i) = ModOrder.CampoManuale(i)
              Text1(i).Text = Cartellino.CampoManuale(i)
           End If
           campo(i).caption = Cartellino.Fisso(i)
            Text1(i).Visible = Cartellino.FissoVisibile(i)
            campo(i).Visible = Cartellino.FissoVisibile(i)
         Next
     '  Cartellino.DrawLabel Video, Picture1
       'Call Update
       Cartellino.Lingua = Ing  'forzatura cartellino inglese
       
       If Cartellino.Lingua > 0 Then
          OptionLingua(Cartellino.Lingua).value = True
       Else
          OptionLingua(1).value = True
       End If
       
    Else
       SSTab1.TabEnabled(1) = False
    End If
    SSTab1.Tab = 0
    ImgProfilo.Visible = ModRecipe.Profilo > 0
    If ModRecipe.Profilo > 0 Then
       ImgProfilo.Picture = LoadPicture("..\Bitmap\" & NomeProfili(ModRecipe.Profilo) & ".gif")
       PaccoModOrdine.Move 480
       TuboModOrdine.Move 4890
    Else
       PaccoModOrdine.Move 1740
       TuboModOrdine.Move 6330
    End If
    '========================================================================
    lblModOrdine.caption = ModRecipe.NumeroTubiPacco
'    LblPesoModOrdine.caption = Round(ModRecipe.TuboPesoTeorico * ModRecipe.NumeroTubiPacco, 2) & IIf(Conv_UM.SI_metrico, " [kg]", " [lb]")
    LblPesoModOrdine.caption = Round(Conv_UM.Conversione(ModRecipe.TuboLunghezza, UM.mt, UM.ft, 0) * ModRecipe.WeightPerFeet * ModRecipe.NumeroTubiPacco, 0) & IIf(Conv_UM.SI_metrico, " [kg]", " [lb]")
    '========================================================================
    oneShot = True
'    Command3.Visible = Not SSTab3.Visible
End Sub

Sub ComModRicetta_Click()
    Dim i As Integer
    
    Ricetta.IDRicetta = ModRecipe.IDRicetta
    OrdersForm.LoadRecipeData Ricetta
    RecipeModifyForm.StrapEnable = True
    RecipeModifyForm.BundleEnable = True
    RecipeModifyForm.Show (vbModal)
    If RecipeModifyForm.OK Then
        ' copia la ricetta della finestra modifica ricetta
        ' nella ricetta della finestra di modifica ordine
        ModRecipe.TipoPacco = Ricetta.TipoPacco
        ModRecipe.TipoTubo = Ricetta.TipoTubo
        For i = 1 To MAX_ROWS
            ModRecipe.TubiFila(i) = Ricetta.TubiFila(i)
        Next i
        ModRecipe.TuboAltezza = Ricetta.TuboAltezza
        ModRecipe.TuboLarghezza = Ricetta.TuboLarghezza
        ModRecipe.TuboLunghezza = Ricetta.TuboLunghezza
        ModRecipe.TuboSpessore = Ricetta.TuboSpessore
        For i = 1 To MAX_STRAPS
            ModRecipe.QuotaReggia(i) = Ricetta.QuotaReggia(i)
        Next i
        ModRecipe.NumeroRegge = Ricetta.NumeroRegge
        ModRecipe.ControlloDatiPacco
        ModRecipe.VelMPS = Ricetta.VelMPS
        ModRecipe.VelVR1 = Ricetta.VelVR1
        ModRecipe.VelVR2 = Ricetta.VelVR2
        ModRecipe.VelMB1 = Ricetta.VelMB1
        ModRecipe.VelMB2 = Ricetta.VelMB2
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
    
    Frame2.Visible = Param.GetBit("Par229_STORAGE_DEST")
    Label8 = ModRecipe.Destination
    
    'TUBO
    TuboModOrdine.aConfig = Param.GetNumber("Par101_MisureMetriche") + 2 + 4 + 8 'mm+tubo+spessore+label
    TuboModOrdine.aTube_Height = Conv_UM.Conversione(ModRecipe.TuboAltezza, UM.mt, UM.inch, 2) * 100
    TuboModOrdine.aTube_Width = Conv_UM.Conversione(ModRecipe.TuboLarghezza, UM.mt, UM.inch, 2) * 100
    TuboModOrdine.aTube_Tickness = Conv_UM.Conversione(ModRecipe.TuboSpessore, UM.mt, UM.inch, 4) * 1000
    TuboModOrdine.aTipoTubo = ModRecipe.TipoTubo
    
    'PACCO
    PaccoModOrdine.aConfig = Param.GetNumber("Par101_MisureMetriche") + 0 + 4 + 8 + Abs(ModRecipe.Profilo > 0) * 16 'mm+pacco+spessore+label
    PaccoModOrdine.aTube_Height = Conv_UM.Conversione(ModRecipe.TuboAltezza, UM.mt, UM.inch, 2) * 100
    PaccoModOrdine.aTube_Width = Conv_UM.Conversione(ModRecipe.TuboLarghezza, UM.mt, UM.inch, 2) * 100
    PaccoModOrdine.aTube_Tickness = Conv_UM.Conversione(ModRecipe.TuboSpessore, UM.mt, UM.inch, 4) * 1000
    PaccoModOrdine.aTipoTubo = ModRecipe.TipoTubo
    PaccoModOrdine.aTipoPacco = ModRecipe.TipoPacco
    PaccoModOrdine.aTipoProfilo = ModRecipe.Profilo
   
    PaccoModOrdine.aCounted = 0
    
    Dim i As Integer
    For i = 1 To MAX_ROWS
        PaccoModOrdine.TubiFila(i) = ModRecipe.TubiFila(i)
    Next
    
    TuboModOrdine.Refresh
    PaccoModOrdine.Refresh

End Sub

Private Sub DisegnoReggeRefresh()
    ControlloRegge.VisualizzaLabelLunghezza = True
    ControlloRegge.PaccoLunghezza = ModRecipe.TuboLunghezza
    ControlloRegge.VisualizzaQuote = Not (Param.GetBit("Par104_Frontale"))
    Dim i As Integer
    For i = 1 To 12
        ControlloRegge.QuotaReggia(i) = Abs(Not (Param.GetBit("Par104_Frontale"))) * ModRecipe.QuotaReggia(i)
    Next
    ControlloRegge.Refresh
End Sub
      
Private Sub DisplayDescrizione_Click()
    TOUCHKeyBoard.Dati = ModOrder.Descrizione
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        DisplayDescrizione.caption = TOUCHKeyBoard.Dati
        ModOrder.Descrizione = TOUCHKeyBoard.Dati
    End If
   ' PulsantePremuto = True
End Sub

'***********************************************************
' Funzioni di risposta ai comandi dell'operatore ok e cancel
'***********************************************************
Private Sub CancelCommand_Click()
    oneShot = False
    OK = False
    Me.Hide
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

Private Sub ListaRicette_Click()
    Dim RicettaPresente As Boolean
    Dim i As Integer
    
    ModRecipe.IDRicetta = ListaRicette.List(ListaRicette.ListIndex)
    ' legge dati ricetta selezionata
    OrdersForm.LoadRecipeData ModRecipe
    ' controlla se la ricetta è presente in macchina
    '=====================================================
     'Disegno pacco-tubo
    DisegnoPaccoTuboRefresh
    'Disegno regge
    DisegnoReggeRefresh
    
    CmdModificaRicetta.Visible = ModRecipe.IDRicetta <> PaginaPacco.Ricetta_Descrizione And ModRecipe.IDRicetta <> PaginaStoccaggio.Ricetta_Descrizione And _
                                                ModRecipe.IDRicetta <> PaginaReggiatura.Ricetta_Descrizione And ModRecipe.IDRicetta <> PaginaEntrata.Ricetta_Descrizione And _
                                                ModRecipe.IDRicetta <> PaginaWb.Ricetta_Descrizione
    CmdCancellaRicetta.Visible = CmdModificaRicetta.Visible
    ImgProfilo.Visible = ModRecipe.Profilo > 0
    If ModRecipe.Profilo > 0 Then
       ImgProfilo.Picture = LoadPicture("..\Bitmap\" & NomeProfili(ModRecipe.Profilo) & ".gif")
       PaccoModOrdine.Move 480
       TuboModOrdine.Move 4890
    Else
       PaccoModOrdine.Move 1740
       TuboModOrdine.Move 6330
    End If
    '========================================================================
    lblModOrdine.caption = ModRecipe.NumeroTubiPacco
'    LblPesoModOrdine.caption = Conv_UM.Conversione(ModRecipe.TuboPesoTeorico * ModRecipe.NumeroTubiPacco, UM.kg, UM.lB, 1) & IIf(Conv_UM.SI_metrico, " [kg]", " [lb]")
    LblPesoModOrdine.caption = Round(Conv_UM.Conversione(ModRecipe.TuboLunghezza, UM.mt, UM.ft, 0) * ModRecipe.WeightPerFeet * ModRecipe.NumeroTubiPacco, 0) & IIf(Conv_UM.SI_metrico, " [kg]", " [lb]")
    '========================================================================
'    ' aggiornamento campi manuali del cartellino con quelli contenuti nella ricetta
    '========================================================================
    ModOrder.CampoManuale(4) = ModRecipe.IDRicetta
''    Ricetta.IDRicetta = ModRecipe.IDRicetta
    ModOrder.CampoManuale(2) = ModRecipe.Grade
'    Ricetta.Grade = ModRecipe.Grade
    ModOrder.CampoManuale(5) = ModRecipe.Pieces
'    Ricetta.Pieces = ModRecipe.Pieces
    ModOrder.CampoManuale(8) = ModRecipe.Itemcode
'    Ricetta.Itemcode = ModRecipe.Itemcode
    ModOrder.CampoManuale(7) = ModRecipe.Weight
'    Ricetta.Weight = ModRecipe.Weight
    Cartellino.UnitMisura = IIf(ModOrder.TicketUnit, TUnitMis.inch, TUnitMis.metrica)
    For i = 1 To 20
       If i < 11 Then
          Cartellino.CampoManuale(i) = ModOrder.CampoManuale(i)
          Text1(i).Text = Cartellino.CampoManuale(i)
       End If
       campo(i).caption = Cartellino.Fisso(i)
       Text1(i).Visible = Cartellino.FissoVisibile(i)
       campo(i).Visible = Cartellino.FissoVisibile(i)
    Next
    
End Sub

Private Sub OkCommand_Click()
Dim i As Integer

        ModOrder.IDRicetta = ListaRicette.List(ListaRicette.ListIndex)
        ModOrder.LinguaCartellino = LinguaSelezionata
        Cartellino.Lingua = LinguaSelezionata
        For i = 1 To 20
           'If i < 11 Then
              ModOrder.CampoManuale(i) = Text1(i).Text
              Cartellino.CampoManuale(i) = Text1(i).Text
           'Else
           '   If SSTab2.TabEnabled(1) = True Then
           '      Cartellino.CampoAuto(i - 10) = Text1(i).Text
           '   End If
           'End If
        Next
        oneShot = False
        OK = True
        Me.Hide
End Sub

Public Sub RefreshListaRicette(Optional ByVal Filtro As String)
    Dim TrovataRicetta As Integer
    Dim i As Integer
    
    On Error Resume Next
    With OrdersForm.AdoRicette
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdText
       .RecordSource = "SELECT * FROM Recipes" & Filtro
        .Refresh
        If .Recordset.EOF Then
           .RecordSource = "SELECT * FROM Recipes"
           .Refresh
        End If
     
        .Recordset.MoveFirst
        ListaRicette.Clear
        While .Recordset.EOF = False
            If (OrdersForm.CmdRicetteEnable = True And OrdersForm.CmdRicetteEnable And (frmKernel.IDOrdineCorrente <> .Recordset.Fields("ID"))) Or OrdersForm.CmdRicetteEnable = False Then
                If .Recordset.Fields("ID") <> "00" Then ListaRicette.AddItem .Recordset.Fields("ID")
            End If
            .Recordset.MoveNext
        Wend
        Set .Recordset.ActiveConnection = Nothing
    End With
    
    Dim Trovato As Boolean
    
    Trovato = False
    For i = 0 To ListaRicette.ListCount - 1
       If ModOrder.IDRicetta = ListaRicette.List(i) Then
          ListaRicette.ListIndex = i
          Trovato = True
        Exit For
       End If
    Next
    If Trovato = False Then ListaRicette.ListIndex = 0
    
End Sub

Private Sub SelettoreSagome_Click()
   ModOrder.TicketUnit = Not ModOrder.TicketUnit
   SelettoreSagome.Picture = ImageList1.ListImages(Abs(ModOrder.TicketUnit) + 1).Picture
   Cartellino.UnitMisura = IIf(Cartellino.UnitMisura = TUnitMis.metrica, TUnitMis.inch, metrica)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim i As Integer
   
   For i = 1 To 20
       campo(i).caption = Cartellino.Fisso(i)
       Text1(i).Visible = Cartellino.FissoVisibile(i)
       campo(i).Visible = Cartellino.FissoVisibile(i)
  Next

End Sub

Private Sub Text1_Click(Index As Integer)
 Select Case Index
 Case 0
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 32000
    TOUCHNumericPad.Dati = Text1(0).Text
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        Text1(0).Text = TOUCHNumericPad.Dati
        ModOrder.PresetPacchi = TOUCHNumericPad.Dati
    End If
 Case 1  'customer input
    TOUCHKeyBoard.TextModifica.PasswordChar = ""
    TOUCHKeyBoard.Dati = Text1(Index).Text
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        Text1(Index).Text = TOUCHKeyBoard.Dati
        If Index < 11 Then
           If Trim(TOUCHKeyBoard.Dati) = "" Then TOUCHKeyBoard.Dati = "STOCK"
           Cartellino.CampoManuale(Index) = Trim(TOUCHKeyBoard.Dati)
           Text1(Index).Text = Cartellino.CampoManuale(Index)
        End If
    End If
    TOUCHKeyBoard.TextModifica.PasswordChar = ""
' Case 3  'customer input
'    TOUCHKeyBoard.TextModifica.PasswordChar = ""
'    TOUCHKeyBoard.Dati = Text1(Index).Text
'    TOUCHKeyBoard.Show vbModal
'    If TOUCHKeyBoard.DatiConfermati Then
'        Text1(Index).Text = TOUCHKeyBoard.Dati
'        If Index < 11 Then
'           If Trim(TOUCHKeyBoard.Dati) = "" Then TOUCHKeyBoard.Dati = "STOCK"
'           Cartellino.CampoManuale(Index) = Trim(TOUCHKeyBoard.Dati)
'        End If
'    End If
'    TOUCHKeyBoard.TextModifica.PasswordChar = ""
 Case Else
    TOUCHKeyBoard.TextModifica.PasswordChar = ""
    TOUCHKeyBoard.Dati = Text1(Index).Text
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        Text1(Index).Text = TOUCHKeyBoard.Dati
        If Index < 11 Then
           Cartellino.CampoManuale(Index) = TOUCHKeyBoard.Dati
           Text1(Index).Text = Cartellino.CampoManuale(Index)
        End If
    End If
    TOUCHKeyBoard.TextModifica.PasswordChar = ""
 End Select
End Sub

'==============================================================================
' scorrimento nella lista
'==============================================================================

'Private Sub Timer1_Timer()
'Dim MouseOK As Boolean
'
'   GetCursorPos lpPoint
'   MouseOK = False
'   If GetAsyncKeyState(vbKeyLButton) < 0 Then MouseOK = True
'   If lpPoint.x > 40 And lpPoint.x < 245 And lpPoint.y > 180 And lpPoint.y < 210 And MouseOK = True Then
'      If ListaRicette.ListIndex > 0 Then
'         ListaRicette.ListIndex = ListaRicette.ListIndex - 1
'      End If
'   End If
'
'   If lpPoint.x > 40 And lpPoint.x < 245 And lpPoint.y > 630 And lpPoint.y < 670 And MouseOK = True Then
'      If ListaRicette.ListIndex < ListaRicette.ListCount - 1 Then
'         ListaRicette.ListIndex = ListaRicette.ListIndex + 1
'      End If
'   End If
'End Sub

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

Sub ScritteMultilingua()
    SSTab1.TabCaption(0) = Param.Text("ORDER")
    FrameOrderEnd.caption = Param.Text("AT ORDER END")
    FrameSetup.caption = Param.Text("SETUP")
    CancelCommand.caption = Param.Text("Annulla")
    Label1.caption = Param.Text("Dopo")
    FrameOrderEnd.caption = Param.Text("Ricette")
    SSTab1.TabCaption(1) = Param.Text("Ticket")
    SSTab2.TabCaption(1) = Param.Text("Automatici")
    SSTab2.TabCaption(0) = Param.Text("Manuali")
    CmdNuovaRicetta.caption = Param.Text("nuovo")
    CmdModificaRicetta.caption = Param.Text("MODIFICA")
    CmdCancellaRicetta.caption = Param.Text("Cancella")
    Frame1.caption = "Recipe operations"
    CloseCommand.caption = Param.Text("Chiudi")
    Label2.caption = Param.Text("Pacchi")
    Combo1.Clear
    Combo1.AddItem Param.Text("Non stop")
    Combo1.AddItem Param.Text("Arresto")
    Combo1.AddItem Param.Text("Nuovo Ordine")
'    Label5.caption = Param.Text("Nticket")
    Label4.caption = Param.Text("Filtro")
    Label7.caption = Param.Text("Ntubi")
    Label6.caption = Param.Text("Peso2")
'    Label5 = Param.Text("Nticket")
End Sub


