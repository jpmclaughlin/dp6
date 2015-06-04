VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form OrdersForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9345
      Left            =   0
      TabIndex        =   16
      Top             =   1050
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   16484
      _Version        =   393216
      TabHeight       =   882
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Next orders"
      TabPicture(0)   =   "OrdersForm.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DisplayDescrizione(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DisplayID(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DisplayDescrizione(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LedPrenotato"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "GridOrdiniFuturi"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ComShiftUpOrdine"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ComShiftDownOrdine"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ComModificaOrdine"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "ComNuovoOrdine"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ComCancellaOrdine"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CmdRicette"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Timerlocale"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Actual orders"
      TabPicture(1)   =   "OrdersForm.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(7)=   "Frame2(0)"
      Tab(1).Control(8)=   "Frame2(1)"
      Tab(1).Control(9)=   "Frame2(2)"
      Tab(1).Control(10)=   "Frame2(3)"
      Tab(1).Control(11)=   "Frame2(4)"
      Tab(1).Control(12)=   "Frame2(5)"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "Old orders"
      TabPicture(2)   =   "OrdersForm.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Timer Timerlocale 
         Interval        =   500
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton CmdRicette 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ricette"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   13020
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   780
         Width           =   2115
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   6510
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Index           =   5
         Left            =   -74940
         TabIndex        =   73
         Top             =   5220
         Width           =   15105
         Begin VB.Label Tubes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   13740
            TabIndex        =   81
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   12690
            TabIndex        =   80
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeT 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   11640
            TabIndex        =   79
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   10590
            TabIndex        =   78
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   9540
            TabIndex        =   77
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "Storage Zone"
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
            Height          =   345
            Index           =   5
            Left            =   90
            TabIndex        =   76
            Top             =   300
            Width           =   2685
         End
         Begin VB.Label DisplayDescrizione 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   3630
            TabIndex        =   75
            Top             =   180
            Width           =   5925
         End
         Begin VB.Label DisplayID 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   2790
            TabIndex        =   74
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Index           =   4
         Left            =   -74940
         TabIndex        =   64
         Top             =   4470
         Width           =   15105
         Begin VB.Label DisplayID 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   2790
            TabIndex        =   72
            Top             =   180
            Width           =   855
         End
         Begin VB.Label DisplayDescrizione 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   3630
            TabIndex        =   71
            Top             =   180
            Width           =   5925
         End
         Begin VB.Label Label4 
            Caption         =   "Strapping Zone"
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
            Height          =   345
            Index           =   4
            Left            =   90
            TabIndex        =   70
            Top             =   300
            Width           =   2685
         End
         Begin VB.Label TubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   9540
            TabIndex        =   69
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   10590
            TabIndex        =   68
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeT 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   11640
            TabIndex        =   67
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   12690
            TabIndex        =   66
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Tubes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   13740
            TabIndex        =   65
            Top             =   180
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Index           =   3
         Left            =   -74940
         TabIndex        =   55
         Top             =   3720
         Width           =   15105
         Begin VB.Label Tubes 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   13740
            TabIndex        =   63
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   12690
            TabIndex        =   62
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeT 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   11640
            TabIndex        =   61
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   10590
            TabIndex        =   60
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   9540
            TabIndex        =   59
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "Packpipe Zone"
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
            Height          =   345
            Index           =   3
            Left            =   90
            TabIndex        =   58
            Top             =   300
            Width           =   2685
         End
         Begin VB.Label DisplayDescrizione 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   3630
            TabIndex        =   57
            Top             =   180
            Width           =   5925
         End
         Begin VB.Label DisplayID 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   2790
            TabIndex        =   56
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Index           =   2
         Left            =   -74940
         TabIndex        =   46
         Top             =   2970
         Width           =   15105
         Begin VB.Label Tubes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   13740
            TabIndex        =   54
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   12690
            TabIndex        =   53
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeT 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   11640
            TabIndex        =   52
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   10590
            TabIndex        =   51
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   9540
            TabIndex        =   50
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "Blowing Zone"
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
            Height          =   345
            Index           =   2
            Left            =   90
            TabIndex        =   49
            Top             =   300
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.Label DisplayDescrizione 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   3630
            TabIndex        =   48
            Top             =   180
            Width           =   5925
         End
         Begin VB.Label DisplayID 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   2790
            TabIndex        =   47
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Index           =   1
         Left            =   -74940
         TabIndex        =   37
         Top             =   2220
         Width           =   15105
         Begin VB.Label DisplayID 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2790
            TabIndex        =   45
            Top             =   180
            Width           =   855
         End
         Begin VB.Label DisplayDescrizione 
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   3630
            TabIndex        =   44
            Top             =   180
            Width           =   5925
         End
         Begin VB.Label Label4 
            Caption         =   "Rollway zone"
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
            Height          =   345
            Index           =   1
            Left            =   90
            TabIndex        =   43
            Top             =   300
            Width           =   2685
         End
         Begin VB.Label TubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   9540
            TabIndex        =   42
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   10590
            TabIndex        =   41
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeT 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   11640
            TabIndex        =   40
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   12690
            TabIndex        =   39
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Tubes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   13740
            TabIndex        =   38
            Top             =   180
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Height          =   765
         Index           =   0
         Left            =   -74940
         TabIndex        =   23
         Top             =   1470
         Width           =   15105
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Height          =   165
            Left            =   2790
            TabIndex        =   94
            Top             =   -30
            Width           =   12015
         End
         Begin VB.Label Tubes 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   13740
            TabIndex        =   36
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeL 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   12690
            TabIndex        =   35
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeT 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   11640
            TabIndex        =   34
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeH 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   10590
            TabIndex        =   33
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label TubeW 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   9540
            TabIndex        =   32
            Top             =   180
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "MILL"
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
            Height          =   345
            Index           =   0
            Left            =   90
            TabIndex        =   31
            Top             =   300
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.Label DisplayDescrizione 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ---"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   3630
            TabIndex        =   27
            Top             =   180
            Width           =   5925
         End
         Begin VB.Label DisplayID 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  ---"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   26
            Top             =   180
            Width           =   855
         End
      End
      Begin VB.CommandButton ComCancellaOrdine 
         BackColor       =   &H00C1C1C1&
         Caption         =   "CANCELLA"
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
         Left            =   7260
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   780
         Width           =   2055
      End
      Begin VB.CommandButton ComNuovoOrdine 
         BackColor       =   &H00C1C1C1&
         Caption         =   "NUOVO"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   780
         Width           =   2055
      End
      Begin VB.CommandButton ComModificaOrdine 
         BackColor       =   &H00C1C1C1&
         Caption         =   "MODIFICA"
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
         Left            =   3030
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton ComShiftDownOrdine 
         BackColor       =   &H00C1C1C1&
         Caption         =   "ß"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   11190
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   1785
      End
      Begin VB.CommandButton ComShiftUpOrdine 
         BackColor       =   &H00C1C1C1&
         Caption         =   "Ý"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   20.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   9360
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   1785
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridOrdiniFuturi 
         Height          =   6600
         Left            =   2160
         TabIndex        =   17
         Top             =   2610
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   11642
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   8421504
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Shape LedPrenotato 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   870
         Shape           =   3  'Circle
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "NEXT orders"
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
         Height          =   495
         Left            =   540
         TabIndex        =   95
         Top             =   5190
         Width           =   1515
      End
      Begin VB.Label DisplayDescrizione 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ---"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   8910
         TabIndex        =   93
         Top             =   2130
         Width           =   5955
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   6555
         Index           =   7
         Left            =   390
         TabIndex        =   92
         Top             =   2640
         Width           =   1755
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Packpipe order"
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
         Height          =   495
         Index           =   6
         Left            =   390
         TabIndex        =   91
         Top             =   2130
         Width           =   1755
      End
      Begin VB.Label DisplayID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  ---"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   2190
         TabIndex        =   89
         Top             =   2130
         Width           =   825
      End
      Begin VB.Label DisplayDescrizione 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ---"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3000
         TabIndex        =   88
         Top             =   2130
         Width           =   5955
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tubes"
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
         Height          =   435
         Left            =   -61200
         TabIndex        =   86
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tube L"
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
         Height          =   435
         Left            =   -62250
         TabIndex        =   85
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tube T"
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
         Height          =   435
         Left            =   -63300
         TabIndex        =   84
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tube H"
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
         Height          =   435
         Left            =   -64350
         TabIndex        =   83
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tube W"
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
         Height          =   435
         Left            =   -65400
         TabIndex        =   82
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID"
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
         Height          =   435
         Left            =   -72150
         TabIndex        =   30
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
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
         Height          =   435
         Left            =   -71340
         TabIndex        =   29
         Top             =   1020
         Width           =   5955
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recipe ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   8910
         TabIndex        =   28
         Top             =   1710
         Width           =   5955
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   3000
         TabIndex        =   25
         Top             =   1710
         Width           =   5925
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2190
         TabIndex        =   24
         Top             =   1710
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   14
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   9
         Top             =   270
         Visible         =   0   'False
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
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   270
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "OrdersForm.frx":0054
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
         TabIndex        =   12
         Top             =   150
         Width           =   8985
      End
   End
   Begin MSAdodcLib.Adodc AdoRicette 
      Height          =   540
      Left            =   9030
      Top             =   5895
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   953
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
      Caption         =   "AdoRicette"
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
   Begin MSAdodcLib.Adodc AdoOrdini 
      Height          =   525
      Left            =   9030
      Top             =   5175
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   926
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\usr\dp6.pc\target\Produzione.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\usr\dp6.pc\target\Produzione.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Ordini where( Visualizzato=true) orderby id asc"
      Caption         =   "AdoOrdini"
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
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EE3959&
      BackStyle       =   0  'Transparent
      Caption         =   "Attuale"
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
      Height          =   435
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   2205
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
      Left            =   12510
      TabIndex        =   1
      Top             =   120
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
      Left            =   12990
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "OrdersForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Ini As Boolean
Private RigaSelezionata As Integer
Private Conferma As Boolean
Private PassformFocus As Boolean
Public CmdRicetteEnable As Boolean
Private oneShot As Boolean

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub CmdRicette_Click()
    Dim IDOrdine As Integer
    
    CmdRicetteEnable = True
    TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
    TechPasswordForm.Show vbModal
    If TechPasswordForm.LoginSucceeded = False Then Exit Sub
    Unload TechPasswordForm
          
    ' se ci sono Orders futuri
    CmdRicetteEnable = True
    If GridOrdiniFuturi.Rows > 0 Then
       ' legge dati ordine selezionato
       GridOrdiniFuturi.Col = 0    ' ID ordine è sulla colonna 0
       ModOrder.IDOrdine = Val(GridOrdiniFuturi.Text)
    Else
       ModOrder.IDOrdine = 1
    End If
        If LoadOrderData(ModOrder) Then
            ModRecipe.IDRicetta = ModOrder.IDRicetta
            ' legge dati ricetta selezionata
            LoadRecipeData ModRecipe
            Conferma = False
            ' chiama finestra di modifica
            OrderModifyForm.SSTab1.Tab = 0
            OrderModifyForm.PulsantePremuto = True
            OrderModifyForm.Show vbModal
            If OrderModifyForm.OK Then
               Conferma = True
            End If
           '8) se dati confermati, li salva
           If Conferma Then
                SaveOrderData ModOrder
                SaveRecipeData ModRecipe
                RefreshListaFuturi
                ' forza il bit di dati non validi sulla commessa futura,
                ' così il kernel la ritrasmette
                DB402.Bit(0, 0) = False
            End If
        End If
End Sub

Private Sub Command1_Click()
    
    GridOrdiniFuturi.Col = 1
    DialogChangeOrder.Prenotato = DB402.Bit(0, 1)
    DialogChangeOrder.TextInt = Trim(GridOrdiniFuturi.Text)
    DialogChangeOrder.Show vbModal
    If DialogChangeOrder.Risposta = False Then Exit Sub
    
    If CodiceFuturo > 0 Then
        DB402.Bit(0, 1) = Not DB402.Bit(0, 1)      ' toggle richiesta cambio commessa
    Else
        DB402.Bit(0, 1) = False
    End If
    AggiornaLed
End Sub

' timer locale aggiornamento pagina

Private Sub TimerLocale_Timer()
  Static One As Boolean
  
  Me.Update
  If One Then RefreshAttuali
  One = Not One
End Sub

Private Sub Form_Activate()
    ' old orders not enable
    ' refresh page
    oneShot = True
    Me.Update
    ' recipe button enable
    If Param.GetBit("Par220_AttivaGestioneRicette") = False Then
       CmdRicette.Visible = False
    Else
       CmdRicette.Visible = True
    End If
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    RecipeModifyForm.SSTab1.TabEnabled(4) = True
    RecipeModifyForm.SSTab2.Tab = 0
    RecipeModifyForm.SSTab1.Tab = 0
    RecipeModifyForm.SSTab2.TabEnabled(1) = False
    ' refresh page data
    RefreshListaFuturi
    RefreshAttuali
    ' abilitazione temporizzatore locale
    TimerLocale.Enabled = True
    TimerLocale.Interval = 500
    ' (disattivare in Form_deactivate)
    Barra21.Selezionato = 1
    SSTab1.Tab = 0
End Sub

Private Sub RefreshAttuali()
    'number of the tubes
    Tubes(3) = DB420.Word(30)
    ' mill data order
  '  TubeW(0) = IIf(DB448.Word(6) > 0, DB448.Word(6) / 10, DB448.Word(10) / 10)
  '  TubeH(0) = DB448.Word(8) / 10
  '  TubeT(0) = DB448.Word(12) / 100
  '  TubeL(0) = DB448.Word(4)
    ' Actual order data refresh
        tubew(1) = DB450.Word(6) / 10 ' RS_produzione.Fields("Larghezza") * 1000
        tubeh(1) = DB450.Word(8) / 10 ' RS_produzione.Fields("Altezza") * 1000
        TubeT(1) = DB450.Word(10) / 10 'RS_produzione.Fields("Spessore") * 1000
        TubeL(1) = DB450.Word(4)  'RS_produzione.Fields("Lunghezza") * 1000
        tubew(3) = DB470.Word(6) / 10 'RS_produzione.Fields("Larghezza") * 1000
        tubeh(3) = DB470.Word(8) / 10 'RS_produzione.Fields("Altezza") * 1000
        TubeT(3) = DB470.Word(10) / 10 'RS_produzione.Fields("Spessore") * 1000
        TubeL(3) = DB470.Word(4)  'RS_produzione.Fields("Lunghezza") * 1000
        tubew(5) = DB486.Word(6) / 10 'RS_produzione.Fields("Larghezza") * 1000
        tubeh(5) = DB486.Word(8) / 10 'RS_produzione.Fields("Altezza") * 1000
        TubeT(5) = DB486.Word(10) / 10 'RS_produzione.Fields("Spessore") * 1000
        TubeL(5) = DB486.Word(4)  'RS_produzione.Fields("Lunghezza") * 1000
        tubew(4) = DB480.Word(6) / 10 'RS_produzione.Fields("Larghezza") * 1000
        tubeh(4) = DB480.Word(8) / 10 'RS_produzione.Fields("Altezza") * 1000
        TubeT(4) = DB480.Word(10) / 10 'RS_produzione.Fields("Spessore") * 1000
        TubeL(4) = DB480.Word(4)  'RS_produzione.Fields("Lunghezza") * 1000
        DisplayID(1) = DB450.Word(0)
        DisplayID(3) = DB470.Word(0)
        DisplayID(5) = DB486.Word(0)
        DisplayID(6) = DisplayID(3)
        DisplayID(4) = DB480.Word(0)
    If OrderChanged = False And oneShot = False Then Exit Sub
    OrdiniMacchina.Client_OrdiniInMacchina_Refresh True
    If OrdiniMacchina.Client_Ordine_find(frmKernel.CodOrdineCorrente.CodEntrata) Then
        DisplayDescrizione(1) = RS_produzione.Fields("Descrizione")
    Else
        DisplayDescrizione(1) = "No recipe associated"
    End If
'    If OrdiniMacchina.Client_Ordine_find(frmKernel.CodOrdineCorrente.CodLav) Then
'        DisplayID(2) = RS_produzione.Fields("ID")
'        DisplayDescrizione(2) = RS_produzione.Fields("Descrizione")
'        tubew(2) = RS_produzione.Fields("Larghezza") * 1000
'        tubeh(2) = RS_produzione.Fields("Altezza") * 1000
'        TubeT(2) = RS_produzione.Fields("Spessore") * 1000
'        TubeL(2) = RS_produzione.Fields("Lunghezza") * 1000
'    End If
    If OrdiniMacchina.Client_Ordine_find(frmKernel.CodOrdineCorrente.CodPacco) Then
        DisplayDescrizione(3) = RS_produzione.Fields("Descrizione")
        DisplayDescrizione(6) = DisplayDescrizione(3)
        DisplayDescrizione(7) = RS_produzione.Fields("IDRicetta")
    Else
        DisplayDescrizione(3) = "No recipe associated"
        DisplayDescrizione(6) = "No recipe associated"
    End If
    If OrdiniMacchina.Client_Ordine_find(frmKernel.CodOrdineCorrente.CodRegge) Then
        DisplayDescrizione(4) = RS_produzione.Fields("Descrizione")
    Else
        DisplayDescrizione(4) = "No recipe associated"
    End If
    If OrdiniMacchina.Client_Ordine_find(frmKernel.CodOrdineCorrente.CodStoccaggio) Then
        DisplayDescrizione(5) = RS_produzione.Fields("Descrizione")
    Else
        DisplayDescrizione(5) = "No recipe associated"
    End If
    ' recipe modify form refresh
    With RecipeModifyForm
        .Command1(0).Enabled = False
        .Command1(1).Enabled = False
        .Command1(2).Enabled = False
        .Command1(3).Enabled = False
    End With
    OrderChanged = False
    oneShot = False
    
End Sub
'***************************************
' gestione richieste da plc
'***************************************
Public Sub Update()
   Dim LastCode As Integer
   
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
   Call AggiornaLed
   
   If DB402.Bit(0, 1) Then
     ComCancellaOrdine.Enabled = False
   ElseIf GridOrdiniFuturi.Rows > 0 Then
      ComCancellaOrdine.Enabled = True
   End If
    
   LastCode = UltimoCodice
   ComNuovoOrdine.Enabled = LastCode > 0
   If GridOrdiniFuturi.Rows > 0 Then
      ComModificaOrdine.Enabled = True
      ComCancellaOrdine.Enabled = True
      ComShiftUpOrdine.Enabled = True
      ComShiftDownOrdine.Enabled = True
      Command1.Visible = True
   Else
      ComModificaOrdine.Enabled = False
      ComCancellaOrdine.Enabled = False
      ComShiftUpOrdine.Enabled = False
      ComShiftDownOrdine.Enabled = False
      Command1.Visible = False
   End If
   
   If RefreshListaOrdini Then RefreshListaFuturi
   
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
End Sub

Private Sub Form_Load()
    PassformFocus = False
    ScritteMultilingua
    WindowState = vbMinimized
    WindowState = 2
    TimerLocale.Enabled = False
End Sub

' modifica dati ordine attuale o futuro (quello con il Focus)
Private Sub ComModificaOrdine_Click()
    ' se ci sono Orders futuri
    
    CmdRicetteEnable = False
    If GridOrdiniFuturi.Rows > 0 Then
        ' legge dati ordine selezionato
        GridOrdiniFuturi.Col = 0    ' ID ordine è sulla colonna 0
        ModOrder.IDOrdine = Val(GridOrdiniFuturi.Text)
        If LoadOrderData(ModOrder) Then
            ModRecipe.IDRicetta = ModOrder.IDRicetta
            ' legge dati ricetta selezionata
            LoadRecipeData ModRecipe
            Conferma = False
            ' chiama finestra di modifica
            '***************OrderModifyForm.SSTab1.Tab = 0
            ' controlla se è abilitata la gestione delle Recipes quindi attiva la pagina Orders oppure la pagina delle ricette
            If Param.GetBit("Par220_AttivaGestioneRicette") = False Then
                RecipeModifyForm.Update
                'frmStampa.VariableLabelUpdate ModOrder
               Call PaginaRicette
               'RecipeModifyForm.Show (vbModal)
               If RecipeModifyForm.OK Then
                  Conferma = True
               End If
           Else
               OrderModifyForm.RefreshListaRicette
               OrderModifyForm.Show (vbModal)
               If OrderModifyForm.OK Then
                  Conferma = True
               End If
           End If
           '8) se dati confermati, li salva
           If Conferma Then
                SaveOrderData ModOrder
                SaveRecipeData ModRecipe
                RefreshListaFuturi
                ' forza il bit di dati non validi sulla commessa futura,
                ' così il kernel la ritrasmette
                DB402.Bit(0, 0) = False
            End If
        End If
    End If
End Sub

Sub PaginaRicette()
    Dim i As Integer
    
    If OrderModifyForm.NewOrModify = False Or Param.GetBit("Par220_AttivaGestioneRicette") = False Then
       Ricetta.IDRicetta = ModRecipe.IDRicetta
       OrdersForm.LoadRecipeData Ricetta
    End If
    RecipeModifyForm.StrapEnable = True
    RecipeModifyForm.BundleEnable = True
    RecipeModifyForm.Show (vbModal)
    If RecipeModifyForm.OK Then
        ' copia la ricetta della finestra modifica ricetta
        ' nella ricetta della finestra di modifica ordine
        ModRecipe.TipoPacco = Ricetta.TipoPacco
        ModRecipe.WeightPerFeet = Ricetta.WeightPerFeet
        ModRecipe.TipoTubo = Ricetta.TipoTubo
        For i = 1 To MAX_ROWS
            ModRecipe.TubiFila(i) = Ricetta.TubiFila(i)
        Next i
        ModRecipe.Destination = Ricetta.Destination
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
        ModRecipe.Bypass0 = Ricetta.Bypass0
        ModRecipe.Bypass1 = Ricetta.Bypass1
        ModRecipe.Bypass2 = Ricetta.Bypass2
        ModRecipe.Bypass3 = Ricetta.Bypass3
        ModRecipe.VelVR1 = Ricetta.VelVR1
        ModRecipe.VelVR2 = Ricetta.VelVR2
        ModRecipe.VelMB1 = Ricetta.VelMB1
        ModRecipe.VelMB2 = Ricetta.VelMB2
        ModRecipe.VelTR = Ricetta.VelTR
        ModRecipe.TipoCalcRegge = Ricetta.TipoCalcRegge
        ModRecipe.Regg1 = Ricetta.Regg1
        ModRecipe.Regg2 = Ricetta.Regg2
        ModRecipe.Profilo = Ricetta.Profilo
    End If

End Sub
 
' Append new order
Private Sub ComNuovoOrdine_Click()
    Dim LastCode As Integer
    
    CmdRicetteEnable = False
    LastCode = UltimoCodice
    '1) verifica se c'è spazio per un nuovo ordine
    If LastCode > 0 Then
        '2) carica in ModOrder l'ultimo ordine programmato
        ModOrder.IDOrdine = LastCode
        LoadOrderData ModOrder
        '3) carica in ModRecipe l'ultima ricetta programmata
        ModRecipe.IDRicetta = ModOrder.IDRicetta
        LoadRecipeData ModRecipe
        '4) incrementa il codice ordine
        ModOrder.IDOrdine = ModOrder.IDOrdine + 1
        If ModOrder.IDOrdine >= 100 Then ModOrder.IDOrdine = 1
        '5) genera il nome ricetta se versione senza ricette: lo genera dal nome ordine
        If Param.GetBit("Par220_AttivaGestioneRicette") = False Then
            ModRecipe.IDRicetta = Format(ModOrder.IDOrdine, "00")
            SaveRecipeData ModRecipe
            ModOrder.IDRicetta = ModRecipe.IDRicetta
        End If
        '6) genera descrizione di default per il nuovo ordine
        ModOrder.Descrizione = "Ord" & ModOrder.IDOrdine
        '7) chiama la OrderModifyForm o la recipemodifyform ********************
        OrderModifyForm.SSTab1.Tab = 0
        Conferma = False
        ' controlla se è abilitata la gestione delle Recipes quindi attiva la pagina Orders oppure la pagina delle ricette
        If Param.GetBit("Par220_AttivaGestioneRicette") = False Then
           RecipeModifyForm.Update
           'frmStampa.VariableLabelUpdate ModOrder
           Call PaginaRicette
           'RecipeModifyForm.Show (vbModal)
           If RecipeModifyForm.OK Then
               Conferma = True
           End If
        Else
           OrderModifyForm.RefreshListaRicette
           OrderModifyForm.Show (vbModal)
            If OrderModifyForm.OK Then
               Conferma = True
            End If
        End If
        '8) se dati confermati, li salva
        If Conferma Then
            ModOrder.Visualizzato = True
            SaveOrderData ModOrder
            SaveRecipeData ModRecipe
            '9) Aggiorna lista ordini
            RefreshListaFuturi
            ' forza il bit di dati non validi sulla commessa futura,
            ' così il kernel la ritrasmette
            DB402.Bit(0, 0) = False
        End If
    End If
End Sub

' delete
Private Sub ComCancellaOrdine_Click()
    Dim IDOrdineSucc As Integer
    Dim IDOrdinePrec As Integer
    Dim CodFut As Integer
    Dim i As Integer
    
    i = 0
    CmdRicetteEnable = False
    ' se ci sono Orders futuri
    If GridOrdiniFuturi.Rows > 0 And DB402.Bit(0, 1) = False Then
        ' legge dati ordine selezionato
        GridOrdiniFuturi.Col = 0    ' ID ordine è sulla colonna 0
        ModOrder.IDOrdine = Val(GridOrdiniFuturi.Text)
        If LoadOrderData(ModOrder) Then
            OrderDeleteForm.OrderLabel.caption = ModOrder.Descrizione
            OrderDeleteForm.RecipeOrOrder = False
            OrderDeleteForm.Show vbModal
            If OrderDeleteForm.OK Then
                IDOrdinePrec = ModOrder.IDOrdine
                Do
                    ' carica l'ordine successivo
                    IDOrdineSucc = IDOrdinePrec + 1
                    If IDOrdineSucc >= 100 Then IDOrdineSucc = 1
                    
                    ModOrder.IDOrdine = IDOrdineSucc
                    LoadOrderData ModOrder
                    ModOrder.IDOrdine = IDOrdinePrec
                    SaveOrderData ModOrder
                    
                    IDOrdinePrec = IDOrdinePrec + 1
                    If IDOrdinePrec >= 100 Then IDOrdinePrec = 1
                    
                    i = i + 1   ' controllo numero massomo di cicli
                    
                Loop While (ModOrder.Visualizzato) And (i < 100)
            End If
        End If
    End If
    
    RefreshListaFuturi
    
    ' ritrasmissione ordine futuro
    ' forza il bit di dati non validi sulla commessa futura,
    ' così il kernel la ritrasmette
    DB402.Bit(0, 0) = False
End Sub


' carica un ordine con tutti i dati prelevati dal database
'NB: l'ordine deve essere inizializzato con un codice PLC che serve
'    a reperire tutti gli altri dati
'    ritorna true se dati validi
'
Public Function LoadOrderData(ClientOrder As OrderClass) As Boolean
    Dim i As Integer
    With AdoOrdini
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Orders"
        .Refresh
        .Recordset.MoveFirst
        .Recordset.Find ("ID=" & ClientOrder.IDOrdine)

        If .Recordset.EOF = False Then
            On Error Resume Next
                ClientOrder.Descrizione = .Recordset.Fields("Descrizione")
                ClientOrder.IDRicetta = .Recordset.Fields("IDRicetta")
                ClientOrder.Visualizzato = .Recordset.Fields("Visualizzato")
                ClientOrder.ModoCambioOrdine = .Recordset.Fields("GestioneAFineOrdine")
                ClientOrder.PresetPacchi = .Recordset.Fields("NumPacchi")
                ClientOrder.LinguaCartellino = .Recordset.Fields("LinguaCartellino")
                ClientOrder.TicketUnit = .Recordset.Fields("Unit")
                For i = 1 To 10
                    ClientOrder.CampoManuale(i) = .Recordset.Fields("Cartellino" & i)
                Next i
            On Error GoTo 0
            LoadOrderData = True
        Else
            LoadOrderData = False
        End If
        .Recordset.ActiveConnection = Nothing
    End With
End Function

Public Sub SaveOrderData(ClientOrder As OrderClass)
    Dim i As Integer
    With AdoOrdini
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Orders"
        .Refresh
        .Recordset.MoveFirst
        .Recordset.Find ("ID=" & ClientOrder.IDOrdine)

        If .Recordset.EOF = False Then
            On Error Resume Next
                .Recordset.Fields("Descrizione") = ClientOrder.Descrizione
                .Recordset.Fields("IDRicetta") = ClientOrder.IDRicetta
                .Recordset.Fields("Visualizzato") = ClientOrder.Visualizzato
                .Recordset.Fields("GestioneAFineOrdine") = ClientOrder.ModoCambioOrdine
                .Recordset.Fields("NumPacchi") = ClientOrder.PresetPacchi
                .Recordset.Fields("LinguaCartellino") = ClientOrder.LinguaCartellino
                .Recordset.Fields("Unit") = ClientOrder.TicketUnit
                For i = 1 To 10
                    .Recordset.Fields("Cartellino" & i) = ClientOrder.CampoManuale(i)
                Next i
                .Recordset.Update
            On Error GoTo 0
        End If
        
        .Recordset.ActiveConnection = Nothing
    End With
End Sub


Public Function LoadRecipeData(ClientRecipe As RecipeClass) As Boolean
    Dim i As Integer
    
    On Error GoTo Errore
    
    If ClientRecipe.IDRicetta = "" Then ClientRecipe.IDRicetta = "00"
    With AdoRicette
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Recipes"
        .Refresh
        .Recordset.MoveFirst
        .Recordset.Find ("ID='" & ClientRecipe.IDRicetta & "'")

cont:   If .Recordset.EOF = False Then
            On Error Resume Next
                ClientRecipe.TipoTubo = .Recordset.Fields("TipoTubo")
                ClientRecipe.TuboAltezza = .Recordset.Fields("Altezza")
                ClientRecipe.TuboLarghezza = .Recordset.Fields("Larghezza")
                ClientRecipe.TuboLunghezza = .Recordset.Fields("Lunghezza")
                ClientRecipe.TuboSpessore = .Recordset.Fields("Spessore")
                ClientRecipe.TuboPesoTeorico = .Recordset.Fields("PesoTeoricoTubo")
                ClientRecipe.TipoPacco = .Recordset.Fields("TipoPacco")
                ClientRecipe.NumeroTubiPacco = .Recordset.Fields("NumeroTubi")
                ClientRecipe.NumeroFile = .Recordset.Fields("NumeroFile")
                ClientRecipe.PaccoLarghezzaBaseEsagono = .Recordset.Fields("LarghezzaBasePacco")
                ClientRecipe.PaccoLarghezza = .Recordset.Fields("LarghezzaFilaMax")
                ClientRecipe.PaccoLarghezzaLatoEsagono = .Recordset.Fields("LarghezzaLatoPaccoEsagono")
                ClientRecipe.PaccoAltezza = .Recordset.Fields("AltezzaPacco")
                ClientRecipe.FilaUscitaControsagoma = .Recordset.Fields("FilaUscitaControsagoma")
                ClientRecipe.PaccoPesoTeorico = .Recordset.Fields("PesoTeoricoPacco")
                ClientRecipe.VelMPS = .Recordset.Fields("R_473_64_VelMagneti")
                ClientRecipe.VelVR1 = .Recordset.Fields("VelVR1")
                ClientRecipe.VelVR2 = .Recordset.Fields("VelVR2")
                ClientRecipe.VelMB1 = .Recordset.Fields("VelMB1")
                ClientRecipe.VelMB2 = .Recordset.Fields("VelMB2")
                ClientRecipe.Bypass0 = .Recordset.Fields("Bypass0")
                ClientRecipe.Bypass1 = .Recordset.Fields("Bypass1")
                ClientRecipe.Bypass2 = .Recordset.Fields("Bypass2")
                ClientRecipe.Bypass3 = .Recordset.Fields("Bypass3")
                ClientRecipe.VelTR = .Recordset.Fields("VelTRSal")
                ClientRecipe.Regg1 = .Recordset.Fields("Regg1_enable")
                ClientRecipe.Regg2 = .Recordset.Fields("Regg2_enable")
                ClientRecipe.Grade = .Recordset.Fields("Grade")
                ClientRecipe.Itemcode = .Recordset.Fields("Itemcode")
                ClientRecipe.Pieces = .Recordset.Fields("Pieces")
                ClientRecipe.Weight = .Recordset.Fields("Weight")
                ClientRecipe.WeightPerFeet = .Recordset.Fields("WeightPerFeet")
                ClientRecipe.Destination = Val(.Recordset.Fields("Storage_destinations"))
                ClientRecipe.TipoCalcRegge = .Recordset.Fields("TipoCalcRegge")
                  If IsNull(.Recordset.Fields("Profilo")) = True Then
                   ClientRecipe.Profilo = 0
                Else
                   ClientRecipe.Profilo = .Recordset.Fields("Profilo") * Abs(Param.GetBit("Par201_AbilitazioneProfili"))
                End If
                For i = 1 To MAX_ROWS
                    ClientRecipe.TubiFila(i) = .Recordset.Fields("Fila" & Format(i, "00"))
                Next i
                ClientRecipe.NumeroRegge = .Recordset.Fields("NumeroRegge")
                For i = 1 To MAX_STRAPS
                    ClientRecipe.QuotaReggia(i) = .Recordset.Fields("Reggia" & Format(i, "00"))
                Next i
            On Error GoTo 0
            LoadRecipeData = True
        Else
            LoadRecipeData = False
        End If
        .Recordset.ActiveConnection = Nothing
    End With
Exit Function
Errore:

MsgBox "Errore: la ricetta " & ClientRecipe.IDRicetta & " non è presente nel database ricette", vbExclamation, "DATAPACK 6.0"

GoTo cont
End Function

Public Sub SaveRecipeData(ClientRecipe As RecipeClass)
    Dim i As Integer
    With AdoRicette
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Recipes"
        .Refresh
        .Recordset.MoveFirst
        
        If ClientRecipe.IDRicetta = "" Then
            ClientRecipe.IDRicetta = "01"   ' nome di default della ricetta
        End If
        On Error Resume Next
        .Recordset.Find ("ID='" & ClientRecipe.IDRicetta & "'")

        If .Recordset.EOF = True Then
            ' non trovata
            .Recordset.AddNew
            .Recordset.Fields("ID") = ClientRecipe.IDRicetta
        End If
        
        On Error Resume Next
            .Recordset.Fields("TipoTubo") = ClientRecipe.TipoTubo
            .Recordset.Fields("Altezza") = ClientRecipe.TuboAltezza
            .Recordset.Fields("Larghezza") = ClientRecipe.TuboLarghezza
            .Recordset.Fields("Lunghezza") = ClientRecipe.TuboLunghezza
            .Recordset.Fields("Spessore") = ClientRecipe.TuboSpessore
            .Recordset.Fields("PesoTeoricoTubo") = ClientRecipe.TuboPesoTeorico
            .Recordset.Fields("TipoPacco") = ClientRecipe.TipoPacco
            .Recordset.Fields("NumeroTubi") = ClientRecipe.NumeroTubiPacco
            .Recordset.Fields("NumeroFile") = ClientRecipe.NumeroFile
            .Recordset.Fields("LarghezzaBasePacco") = ClientRecipe.PaccoLarghezzaBaseEsagono
            .Recordset.Fields("LarghezzaFilaMax") = ClientRecipe.PaccoLarghezza
            .Recordset.Fields("LarghezzaLatoPaccoEsagono") = ClientRecipe.PaccoLarghezzaLatoEsagono
            .Recordset.Fields("AltezzaPacco") = ClientRecipe.PaccoAltezza
            .Recordset.Fields("FilaUscitaControsagoma") = ClientRecipe.FilaUscitaControsagoma
            .Recordset.Fields("PesoTeoricoPacco") = ClientRecipe.PaccoPesoTeorico
            .Recordset.Fields("R_473_64_VelMagneti") = ClientRecipe.VelMPS
            .Recordset.Fields("VelVR1") = ClientRecipe.VelVR1
            .Recordset.Fields("VelVR2") = ClientRecipe.VelVR2
            .Recordset.Fields("VelMB1") = ClientRecipe.VelMB1
            .Recordset.Fields("VelMB2") = ClientRecipe.VelMB2
            .Recordset.Fields("Bypass0") = ClientRecipe.Bypass0
            .Recordset.Fields("Bypass1") = ClientRecipe.Bypass1
            .Recordset.Fields("Bypass2") = ClientRecipe.Bypass2
            .Recordset.Fields("Bypass3") = ClientRecipe.Bypass3
            .Recordset.Fields("VelTRSal") = ClientRecipe.VelTR
            .Recordset.Fields("Storage_destinations") = ClientRecipe.Destination
            .Recordset.Fields("Regg1_enable") = ClientRecipe.Regg1
            .Recordset.Fields("Regg2_enable") = ClientRecipe.Regg2
            .Recordset.Fields("TipoCalcRegge") = ClientRecipe.TipoCalcRegge
            .Recordset.Fields("Profilo") = ClientRecipe.Profilo * Abs(Param.GetBit("Par201_AbilitazioneProfili"))
            .Recordset.Fields("Grade") = ClientRecipe.Grade
            .Recordset.Fields("Itemcode") = ClientRecipe.Itemcode
            .Recordset.Fields("Pieces") = ClientRecipe.Pieces
            .Recordset.Fields("Weight") = ClientRecipe.Weight
            .Recordset.Fields("WeightPerFeet") = ClientRecipe.WeightPerFeet
           For i = 1 To MAX_ROWS
                .Recordset.Fields("Fila" & Format(i, "00")) = ClientRecipe.TubiFila(i)
            Next i
            .Recordset.Fields("NumeroRegge") = ClientRecipe.NumeroRegge
            For i = 1 To MAX_STRAPS
                .Recordset.Fields("Reggia" & Format(i, "00")) = ClientRecipe.QuotaReggia(i)
            Next i
            .Recordset.Update

        On Error GoTo 0
        .Recordset.ActiveConnection = Nothing
    End With
End Sub

' restituisce l'ultimo codice presente sulla lista dei futuri
' oppure quello in lavoro se la lista è vuota.
' Se la lista futura è piena , restituisce 0
Private Function UltimoCodice() As Integer
    Dim Trovato As Boolean
    Dim i As Integer
    Dim PrimoCodice As Integer
    Dim Lunghezza As Integer
    
    ' verifica l'ultimo codice presente sulla lista dei futuri
    With AdoOrdini
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Orders"
        .Refresh
        .Recordset.MoveFirst
        i = 0
        UltimoCodice = -1
        While (i < 100) And (UltimoCodice < 0)
            i = i + 1
            If Trovato Then
                If .Recordset.Fields("Visualizzato") = False Then
                    UltimoCodice = i
                Else
                    If i = 99 Then UltimoCodice = 100
                End If
            Else
                If .Recordset.Fields("Visualizzato") = True Then
                    Trovato = True
                    If i = 99 Then UltimoCodice = 100
                End If
            End If
            If (i < 99) Then .Recordset.MoveNext
        Wend
        .Recordset.ActiveConnection = Nothing
    End With
    If UltimoCodice > 0 Then
        UltimoCodice = UltimoCodice - 1
        If UltimoCodice <= 0 Then UltimoCodice = 99
    Else
        UltimoCodice = 0
    End If
    
    
    ' se c'è un ultimo codice, allora verifica la lunghezza della lista
    If UltimoCodice > 0 Then
        PrimoCodice = CodiceFuturo
        Lunghezza = UltimoCodice - PrimoCodice + 1
        If Lunghezza < 0 Then Lunghezza = Lunghezza + 99
    Else
        ' non ci sono commesse future, pertanto la prossima commessa è
        ' futura a quella di entrata DB450
        Lunghezza = 0
        UltimoCodice = DB450.Word(0)
        If UltimoCodice < 1 Then UltimoCodice = 1
        If UltimoCodice > 99 Then UltimoCodice = 99
    End If
    
    ' ultimo controllo : se la lista consente inserimenti allora restituisce
    ' il codice , altrimenti lo azzera
    If Lunghezza > 40 Then
        UltimoCodice = 0
    End If
    
End Function


' restituisce il primo codice con il campo "Visualizzato"=true (è un numero da 1 a 99)
' se non ce ne sono allora restituisce 0
Public Function CodiceFuturo() As Integer
    Dim Trovato As Boolean
    Dim i As Integer
    With AdoOrdini
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Orders"
        .Refresh
        .Recordset.MoveLast
        i = 100
        CodiceFuturo = -1
        While (i > 0) And (CodiceFuturo < 0)
            i = i - 1
            If Trovato Then
                If .Recordset.Fields("Visualizzato") = False Then
                    CodiceFuturo = i
                Else
                    If i = 1 Then CodiceFuturo = 0
                End If
            Else
                If .Recordset.Fields("Visualizzato") = True Then
                    Trovato = True
                    If i = 1 Then CodiceFuturo = 0
                End If
            End If
            If i > 1 Then .Recordset.MovePrevious
        Wend
        .Recordset.ActiveConnection = Nothing
    End With
    CodiceFuturo = CodiceFuturo + 1
    If CodiceFuturo >= 100 Then CodiceFuturo = 1
End Function

' visualizza tutti quelli con il campo "Visualizzato"=true

Public Sub RefreshListaFuturi()
    Dim IDOrdine As Integer
    Dim IndiceRiga As Integer
    
    IDOrdine = CodiceFuturo
    With AdoOrdini
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Orders"
        .Refresh
        .Recordset.Find ("ID=" & IDOrdine)
         
        IndiceRiga = 0
        Command1.Enabled = False
        '=========================== popola lista futuri ==========
        If .Recordset.EOF = False Then
            While .Recordset.Fields("Visualizzato") = True
                If GridOrdiniFuturi.Rows < (IndiceRiga + 1) Then GridOrdiniFuturi.AddItem ""
    
                GridOrdiniFuturi.Col = 0
                GridOrdiniFuturi.Row = IndiceRiga
                GridOrdiniFuturi.RowHeight(IndiceRiga) = 600
                GridOrdiniFuturi.CellAlignment = flexAlignLeftCenter
                GridOrdiniFuturi.Text = .Recordset.Fields("ID")
                
                GridOrdiniFuturi.Col = 1
                GridOrdiniFuturi.CellAlignment = flexAlignLeftCenter
                GridOrdiniFuturi.Text = .Recordset.Fields("Descrizione")
                
                GridOrdiniFuturi.Col = 2
                GridOrdiniFuturi.CellAlignment = flexAlignLeftCenter
                GridOrdiniFuturi.Text = .Recordset.Fields("IDRicetta")
                
                IndiceRiga = IndiceRiga + 1
                .Recordset.MoveNext
                If .Recordset.EOF Then
                    .Recordset.MoveFirst
                End If
                'Label4.Enabled = True
                Command1.Enabled = True
            Wend
        End If
        .Recordset.ActiveConnection = Nothing
    End With
    ' ========================== Larghezza campi ==============
    GridOrdiniFuturi.ColWidth(0) = 800
    GridOrdiniFuturi.ColWidth(1) = 5950
    GridOrdiniFuturi.ColWidth(2) = 5950
    ' ========================== riga selezionata ==============
    GridOrdiniFuturi.Rows = IndiceRiga
    If RigaSelezionata > (IndiceRiga - 1) Then RigaSelezionata = IndiceRiga - 1
    If RigaSelezionata < 0 Then RigaSelezionata = 0
    On Error Resume Next
    GridOrdiniFuturi.Row = RigaSelezionata
    GridOrdiniFuturi.Col = 0
    GridOrdiniFuturi.RowSel = RigaSelezionata
    GridOrdiniFuturi.ColSel = 1
    On Error GoTo 0
    RefreshListaOrdini = False
    DoEvents
    AggiornaLed
End Sub

' registrazione della esecuzione del cambio commessa
' cancellando il campo "Visualizzato" dall'ordine stesso
Public Sub OrdineTrasmesso(Codice As Integer)
    Dim Ordine As OrderClass
    
    Set Ordine = New OrderClass
    Ordine.IDOrdine = Codice
    Ordine.UploadData Codice
    Ordine.Visualizzato = False
    Ordine.DownloadData Codice
    'doevents
    RefreshListaFuturi
    ' aggiorna ordine attuale
    frmKernel.UpdateOrdineCorrente
End Sub


' led cambio commessa prenotato
Private Sub AggiornaLed() 'DB402.Bit(0, 1)
    Static Lamp As Boolean
    
    Lamp = Not Lamp
    If DB402.Bit(0, 1) = True Or DB402.Bit(0, 2) = False Then
        If DB402.Bit(0, 1) Then
           If Lamp Then
              LedPrenotato.BackColor = &HFFFF&
           Else
              LedPrenotato.BackColor = RGB(255, 0, 0)
           End If
        Else
           LedPrenotato.BackColor = RGB(255, 0, 0)
        End If
    Else
      If Command1.Enabled = False Then
        LedPrenotato.BackColor = RGB(210, 210, 210)
     Else
        LedPrenotato.BackColor = &HFF00&
     End If
    End If
End Sub

Private Sub ComShiftUpOrdine_Click()
    Dim TmpOrd1 As OrderClass
    Dim Trovato1 As Boolean
    Dim TmpOrd2 As OrderClass
    Dim Trovato2 As Boolean
    Dim IDOrdine As Integer
    Set TmpOrd1 = New OrderClass
    Set TmpOrd2 = New OrderClass
    
    CmdRicetteEnable = False
    ' se ci sono Orders futuri
    If GridOrdiniFuturi.Rows > 1 Then
        ' legge dati ordine selezionato
        GridOrdiniFuturi.Col = 0    ' ID ordine è sulla colonna 0
        TmpOrd1.IDOrdine = Val(GridOrdiniFuturi.Text)
        TmpOrd2.IDOrdine = TmpOrd1.IDOrdine - 1
        If TmpOrd2.IDOrdine < 1 Then TmpOrd2.IDOrdine = 99
        ' carica l'ordine selezionato ed il suo precedente
        Trovato1 = LoadOrderData(TmpOrd1)
        Trovato2 = LoadOrderData(TmpOrd2)
        ' se i due Orders esistono e sono entrambi sulla lista dei futuri
        ' allora si effettua lo scambio
        If Trovato1 And Trovato2 And TmpOrd1.Visualizzato And TmpOrd2.Visualizzato Then
            IDOrdine = TmpOrd1.IDOrdine
            TmpOrd1.IDOrdine = TmpOrd2.IDOrdine
            TmpOrd2.IDOrdine = IDOrdine
            ' salvo i 2 Orders scambiati
            SaveOrderData TmpOrd1
            SaveOrderData TmpOrd2
            ' la selezione deve seguire l'ordine
            RigaSelezionata = RigaSelezionata - 1
            ' forza il bit di dati non validi sulla commessa futura,
            ' così il kernel la ritrasmette
            DB402.Bit(0, 0) = False
        End If
    End If
    RefreshListaFuturi
End Sub

Private Sub ComShiftDownOrdine_Click()
    Dim TmpOrd1 As OrderClass
    Dim Trovato1 As Boolean
    Dim TmpOrd2 As OrderClass
    Dim Trovato2 As Boolean
    Dim IDOrdine As Integer
    Set TmpOrd1 = New OrderClass
    Set TmpOrd2 = New OrderClass
    
    CmdRicetteEnable = False
    ' se ci sono Orders futuri
    If GridOrdiniFuturi.Rows > 1 Then
        ' legge dati ordine selezionato
        GridOrdiniFuturi.Col = 0    ' ID ordine è sulla colonna 0
        TmpOrd1.IDOrdine = Val(GridOrdiniFuturi.Text)
        TmpOrd2.IDOrdine = TmpOrd1.IDOrdine + 1
        If TmpOrd2.IDOrdine > 99 Then TmpOrd2.IDOrdine = 1
        ' carica l'ordine selezionato ed il suo precedente
        Trovato1 = LoadOrderData(TmpOrd1)
        Trovato2 = LoadOrderData(TmpOrd2)
        ' se i due Orders esistono e sono entrambi sulla lista dei futuri
        ' allora si effettua lo scambio
        If Trovato1 And Trovato2 And TmpOrd1.Visualizzato And TmpOrd2.Visualizzato Then
            IDOrdine = TmpOrd1.IDOrdine
            TmpOrd1.IDOrdine = TmpOrd2.IDOrdine
            TmpOrd2.IDOrdine = IDOrdine
            ' salvo i 2 Orders scambiati
            SaveOrderData TmpOrd1
            SaveOrderData TmpOrd2
            ' la selezione deve seguire l'ordine
            RigaSelezionata = RigaSelezionata + 1
            ' forza il bit di dati non validi sulla commessa futura,
            ' così il kernel la ritrasmette
            DB402.Bit(0, 0) = False
        End If
    End If
    RefreshListaFuturi
End Sub

Private Sub GridOrdiniFuturi_Click()
    RigaSelezionata = GridOrdiniFuturi.Row
End Sub

Sub ScritteMultilingua()
    Label1.caption = Param.Text("Attuale")
 '   Label2.Caption = Param.Text("Stato")
    Command1.caption = Param.Text("Cambio Ordine")
    SSTab1.TabCaption(0) = Param.Text("Ordini futuri")
    ComModificaOrdine.caption = Param.Text("MODIFICA")
    ComNuovoOrdine.caption = Param.Text("NUOVO")
    ComCancellaOrdine.caption = Param.Text("CANCELLA")
    CmdRicette.caption = Param.Text("Ricette")
    lblbar(5) = Param.Text("Orders page")
    lblbar(1) = Param.Text("Pagina")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
    Label4(6) = Param.Text("000000078")
    Label14 = Param.Text("Ordini futuri")
    Label5 = Param.Text("Descrizione")
    Label6 = Param.Text("000000040")
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
                .NomeFile = "Ordini_pagina.htm"
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
