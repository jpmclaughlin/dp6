VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Param 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerLocale 
      Left            =   8040
      Top             =   9900
   End
   Begin MSAdodcLib.Adodc AdoNumeri 
      Height          =   495
      Left            =   4020
      Top             =   9030
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      Caption         =   "AdoNumeri"
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
   Begin MSAdodcLib.Adodc AdoTesti 
      Height          =   495
      Left            =   4920
      Top             =   9720
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   873
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
      Connect         =   $"Param.frx":0000
      OLEDBString     =   $"Param.frx":0093
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TestiMultilingua"
      Caption         =   "AdoTesti"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   11475
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   16815
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0C0&
         Height          =   1695
         Left            =   4590
         TabIndex        =   63
         Top             =   7620
         Width           =   5565
         Begin dp6.XPButton XPButton1 
            Height          =   1185
            Index           =   4
            Left            =   1800
            TabIndex        =   64
            Top             =   300
            Visible         =   0   'False
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   2090
            TxtText         =   "Presets"
            TxtTop          =   50
            TxtLeft         =   55
            BTYPE           =   3
            IMGTOP          =   5
            IMGLEFT         =   5
            ICONA           =   "..\bitmap\icone\RSEdsUI72.ico"
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
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Height          =   1695
         Left            =   10320
         TabIndex        =   58
         Top             =   7620
         Width           =   4875
         Begin dp6.XPButton XPButton1 
            Height          =   1185
            Index           =   5
            Left            =   2550
            TabIndex        =   59
            Top             =   330
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   2090
            TxtText         =   "Parametri dp6"
            TxtTop          =   50
            TxtLeft         =   25
            BTYPE           =   3
            IMGTOP          =   5
            IMGLEFT         =   5
            ICONA           =   "..\bitmap\icone\APIMonitor3.ico"
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
            Height          =   1185
            Index           =   7
            Left            =   90
            TabIndex        =   81
            Top             =   330
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   2090
            TxtText         =   "Electric sheets"
            TxtTop          =   50
            TxtLeft         =   30
            BTYPE           =   3
            IMGTOP          =   5
            IMGLEFT         =   5
            ICONA           =   "..\bitmap\icone\MEMacroEditor0.ico"
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
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Height          =   2535
         Left            =   10320
         TabIndex        =   55
         Top             =   5070
         Width           =   4905
         Begin VB.CommandButton Command1 
            Caption         =   "SET"
            Height          =   435
            Left            =   3870
            TabIndex        =   61
            Top             =   750
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
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
            Left            =   90
            TabIndex        =   56
            Text            =   "Combo1"
            Top             =   750
            Width           =   3675
         End
         Begin dp6.XPButton XPButton2 
            Height          =   1065
            Left            =   690
            TabIndex        =   57
            Top             =   1320
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   1879
            TxtText         =   "Ticket layout"
            TxtTop          =   50
            TxtLeft         =   60
            BTYPE           =   3
            IMGTOP          =   5
            IMGLEFT         =   5
            ICONA           =   "..\bitmap\icone\POWERPNT3.ico"
            ImgW            =   10
            ImgH            =   10
            ImgAllarga      =   0   'False
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Printer"
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
            Height          =   405
            Left            =   90
            TabIndex        =   62
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Height          =   405
            Left            =   1650
            TabIndex        =   60
            Top             =   240
            Width           =   3045
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Height          =   3435
         Left            =   10320
         TabIndex        =   46
         Top             =   1590
         Width           =   4905
         Begin VB.Label LabelData 
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
            Height          =   435
            Left            =   180
            TabIndex        =   54
            Top             =   1950
            Width           =   1455
         End
         Begin VB.Label LabelGiorno 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1680
            TabIndex        =   53
            Top             =   1920
            Width           =   825
         End
         Begin VB.Label LabelOre 
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
            Height          =   465
            Left            =   180
            TabIndex        =   52
            Top             =   2730
            Width           =   1425
         End
         Begin VB.Label LabelOra 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1650
            TabIndex        =   51
            Top             =   2700
            Width           =   825
         End
         Begin VB.Label LabelMese 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2730
            TabIndex        =   50
            Top             =   1920
            Width           =   825
         End
         Begin VB.Label LabelAnno 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3750
            TabIndex        =   49
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label LabelMinuti 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2700
            TabIndex        =   48
            Top             =   2700
            Width           =   825
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Modifica data / ora"
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
            Height          =   675
            Left            =   150
            TabIndex        =   47
            Top             =   390
            Width           =   4575
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Height          =   6015
         Left            =   4590
         TabIndex        =   29
         Top             =   1590
         Width           =   5565
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Turno 1"
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
            Left            =   360
            TabIndex        =   45
            Top             =   1500
            Width           =   2325
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Turno 2"
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
            Left            =   330
            TabIndex        =   44
            Top             =   3090
            Width           =   2325
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Turno 3"
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
            Left            =   360
            TabIndex        =   43
            Top             =   4650
            Width           =   2325
         End
         Begin VB.Label LblTurnoInizio 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   2790
            TabIndex        =   42
            Top             =   1500
            Width           =   825
         End
         Begin VB.Label LblTurnoInizio 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   2
            Left            =   2790
            TabIndex        =   41
            Top             =   3060
            Width           =   825
         End
         Begin VB.Label LblTurnoInizio 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   2790
            TabIndex        =   40
            Top             =   4620
            Width           =   825
         End
         Begin VB.Label LblTurnoFine 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   4230
            TabIndex        =   39
            Top             =   1500
            Width           =   825
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   0
            Left            =   3750
            TabIndex        =   38
            Top             =   1320
            Width           =   465
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   1
            Left            =   3750
            TabIndex        =   37
            Top             =   2880
            Width           =   465
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   2
            Left            =   3750
            TabIndex        =   36
            Top             =   4440
            Width           =   465
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Definizione turni"
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
            Height          =   675
            Left            =   360
            TabIndex        =   35
            Top             =   390
            Width           =   4875
         End
         Begin VB.Label LblTurnoFine 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   2
            Left            =   4230
            TabIndex        =   34
            Top             =   3060
            Width           =   825
         End
         Begin VB.Label LblTurnoFine 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   4230
            TabIndex        =   33
            Top             =   4620
            Width           =   825
         End
         Begin VB.Label LblAliasTurno 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   2790
            TabIndex        =   32
            Top             =   2280
            Width           =   2325
         End
         Begin VB.Label LblAliasTurno 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   2
            Left            =   2790
            TabIndex        =   31
            Top             =   3840
            Width           =   2325
         End
         Begin VB.Label LblAliasTurno 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   3
            Left            =   2790
            TabIndex        =   30
            Top             =   5340
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   7695
         Left            =   90
         TabIndex        =   21
         Top             =   1590
         Width           =   4305
         Begin VB.OptionButton OptionLingua 
            BackColor       =   &H009A9A9A&
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
            Index           =   6
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   6750
            Width           =   2535
         End
         Begin VB.OptionButton OptionLingua 
            BackColor       =   &H009A9A9A&
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
            Index           =   5
            Left            =   1170
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   5700
            Width           =   2535
         End
         Begin VB.OptionButton OptionLingua 
            BackColor       =   &H009A9A9A&
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
            Index           =   4
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   3540
            Width           =   2535
         End
         Begin VB.OptionButton OptionLingua 
            BackColor       =   &H009A9A9A&
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
            Index           =   3
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2460
            Width           =   2535
         End
         Begin VB.OptionButton OptionLingua 
            BackColor       =   &H009A9A9A&
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
            Index           =   2
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4620
            Width           =   2535
         End
         Begin VB.OptionButton OptionLingua 
            BackColor       =   &H009A9A9A&
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
            Index           =   1
            Left            =   1140
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1380
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Lingua"
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
            Height          =   675
            Left            =   330
            TabIndex        =   28
            Top             =   360
            Width           =   3615
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   450
            Picture         =   "Param.frx":0126
            Top             =   1380
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   450
            Picture         =   "Param.frx":0568
            Top             =   2460
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   420
            Picture         =   "Param.frx":09AA
            Top             =   3330
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   420
            Picture         =   "Param.frx":0DEC
            Top             =   4440
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   420
            Picture         =   "Param.frx":122E
            Top             =   3690
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   420
            Picture         =   "Param.frx":1670
            Top             =   4800
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   420
            Picture         =   "Param.frx":1AB2
            Top             =   5700
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   450
            Picture         =   "Param.frx":1EF4
            Top             =   6780
            Width           =   480
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   915
            Index           =   1
            Left            =   330
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   915
            Index           =   3
            Left            =   330
            Top             =   2220
            Width           =   3615
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   915
            Index           =   4
            Left            =   330
            Top             =   3300
            Width           =   3615
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   915
            Index           =   2
            Left            =   330
            Top             =   4380
            Width           =   3615
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   915
            Index           =   5
            Left            =   330
            Top             =   5430
            Width           =   3615
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   915
            Index           =   6
            Left            =   330
            Top             =   6540
            Width           =   3615
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   1095
         Left            =   0
         TabIndex        =   9
         Top             =   -60
         Width           =   15375
         Begin dp6.XPButton XPButton1 
            Height          =   885
            Index           =   2
            Left            =   12270
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   80
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
         Begin VB.Image Image4 
            Height          =   1050
            Left            =   3180
            Picture         =   "Param.frx":21FE
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
            TabIndex        =   18
            Top             =   270
            Visible         =   0   'False
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
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
            TabIndex        =   15
            Top             =   270
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Line Line2 
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
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
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
            TabIndex        =   13
            Top             =   630
            Width           =   3495
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   885
            Index           =   2
            Left            =   3180
            TabIndex        =   19
            Top             =   150
            Width           =   8985
         End
      End
      Begin dp6.Barra2 Barra21 
         Height          =   1215
         Left            =   0
         TabIndex        =   20
         Top             =   10410
         Width           =   15405
         _ExtentX        =   27173
         _ExtentY        =   2037
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7605
         Left            =   120
         TabIndex        =   65
         Top             =   1680
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   13414
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   882
         BackColor       =   12632256
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Encoder"
         TabPicture(0)   =   "Param.frx":428C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label16"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label17"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label18"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label19"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label10"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Check1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "FrameLav"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "----"
         TabPicture(1)   =   "Param.frx":42A8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         Begin VB.Frame FrameLav 
            Height          =   795
            Left            =   5400
            TabIndex        =   73
            Top             =   1080
            Width           =   9585
            Begin VB.CommandButton ComdPreset 
               BackColor       =   &H0000FFFF&
               Caption         =   "SET"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   8070
               Style           =   1  'Graphical
               TabIndex        =   74
               Top             =   180
               Width           =   1425
            End
            Begin VB.Label actual 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1530
               TabIndex        =   79
               Top             =   180
               Width           =   1365
            End
            Begin VB.Label Preset 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6270
               TabIndex        =   78
               Top             =   210
               Width           =   1635
            End
            Begin VB.Label Minsoft 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   60
               TabIndex        =   77
               Top             =   180
               Width           =   1365
            End
            Begin VB.Label Maxsoft 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3000
               TabIndex        =   76
               Top             =   210
               Width           =   1365
            End
            Begin VB.Label Offset 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4620
               TabIndex        =   75
               Top             =   210
               Width           =   1365
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enable"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   4170
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   1260
            Width           =   1125
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
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
            Left            =   10335
            TabIndex        =   71
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Soft. fc max"
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
            Left            =   8385
            TabIndex        =   70
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Soft. fc min"
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
            Left            =   5445
            TabIndex        =   69
            Top             =   810
            Width           =   1305
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Preset"
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
            Left            =   12075
            TabIndex        =   68
            Top             =   810
            Width           =   765
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Actual"
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
            Left            =   7260
            TabIndex        =   67
            Top             =   810
            Width           =   735
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Scharf breacker position"
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
            Height          =   465
            Left            =   120
            TabIndex        =   66
            Top             =   1290
            Width           =   3885
         End
      End
   End
   Begin VB.Frame FrameGrid 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   11475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.CommandButton CommandChiudi 
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
         Height          =   795
         Left            =   12540
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   10500
         Width           =   2385
      End
      Begin VB.CommandButton CommandModifica 
         BackColor       =   &H0000FF00&
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
         Height          =   795
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   10470
         Width           =   2205
      End
      Begin VB.CommandButton CommandUp 
         BackColor       =   &H00C0C0C0&
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
         Height          =   735
         Left            =   7140
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   10500
         Width           =   2205
      End
      Begin VB.CommandButton CommandDown 
         BackColor       =   &H00C0C0C0&
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
         Height          =   735
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   10500
         Width           =   2205
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridParametri 
         Height          =   9435
         Left            =   600
         TabIndex        =   1
         Top             =   780
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   16642
         _Version        =   393216
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin dp6.VerticalMenu VerticalMenu1 
         Height          =   9735
         Left            =   12600
         TabIndex        =   6
         Top             =   300
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   17171
         MenusMax        =   5
         MenuCaption1    =   "Utility"
         MenuItemsMax1   =   2
         MenuItemIcon11  =   "Param.frx":42C4
         MenuItemCaption11=   "Database editor"
         MenuItemIcon12  =   "Param.frx":449E
         MenuItemCaption12=   "Help editor"
         MenuCaption2    =   "I/O PC - PLC"
         MenuItemsMax2   =   2
         MenuItemIcon21  =   "Param.frx":47B8
         MenuItemCaption21=   "Download data to PLC"
         MenuItemIcon22  =   "Param.frx":4992
         MenuItemCaption22=   "Upload data from PLC"
         MenuCaption3    =   "Helps"
         MenuItemsMax3   =   2
         MenuItemIcon31  =   "Param.frx":4CAC
         MenuItemCaption31=   "ComS7"
         MenuItemIcon32  =   "Param.frx":4FC6
         MenuItemCaption32=   "Pg / Pc"
         MenuCaption4    =   "Win service"
         MenuItemsMax4   =   4
         MenuItemIcon41  =   "Param.frx":52E0
         MenuItemCaption41=   "Esplora risorse"
         MenuItemIcon42  =   "Param.frx":55FA
         MenuItemCaption42=   "Calcolatrice"
         MenuItemIcon43  =   "Param.frx":5914
         MenuItemCaption43=   "Notepad"
         MenuItemIcon44  =   "Param.frx":5AEE
         MenuItemCaption44=   "mspaint"
         MenuCaption5    =   "DP 6.0"
         MenuItemsMax5   =   2
         MenuItemIcon51  =   "Param.frx":5E08
         MenuItemCaption51=   "Avvia RSM"
         MenuItemIcon52  =   "Param.frx":6122
         MenuItemCaption52=   "Exit program"
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   1065
         Left            =   570
         Shape           =   4  'Rounded Rectangle
         Top             =   10320
         Width           =   11385
      End
      Begin VB.Label LblHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   660
         TabIndex        =   7
         Top             =   120
         Width           =   11295
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   10095
         Left            =   12450
         Shape           =   4  'Rounded Rectangle
         Top             =   150
         Width           =   2505
      End
   End
End
Attribute VB_Name = "Param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public One As Boolean
Private Const TagLingua As String = "Par100_Lingua"
Private IndiceRiga As Long
Public ChiamataService As Boolean
Private ComPesa As Integer
Private NumPagAlarmLog As Integer
Private HelpParam() As String
Private Sub LeggiTabelle()
    Dim Lingua As Integer
     
    With AdoNumeri
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Parameters.mdb;Persist Security Info=False"
        .CommandType = adCmdTable
        .RecordSource = "Parametri"
        
        ' apertura database con eventuale ripristino del file danneggiato
        On Error Resume Next
        .Refresh
        .Recordset.MoveFirst
        .Recordset.Find ("ID='" & TagLingua & "'")
        If .Recordset.EOF = False Then
            Lingua = .Recordset.Fields("Valore")
            If Lingua < 1 Then Lingua = 1
            If Lingua > 6 Then Lingua = 6
        Else
            Lingua = 1
        End If
        
        ' apertura tabella dei testi selezionata
        AdoTesti.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Texts.mdb;Persist Security Info=False"
        AdoTesti.CommandType = adCmdText
        Select Case Lingua
            Case 1
                AdoTesti.RecordSource = "SELECT TagName , ITALIANO AS Valore FROM TestiMultilingua"
            Case 2
                AdoTesti.RecordSource = "SELECT TagName , INGLESE AS Valore FROM TestiMultilingua"
            Case 3
                AdoTesti.RecordSource = "SELECT TagName , FRANCESE AS Valore FROM TestiMultilingua"
            Case 4
                AdoTesti.RecordSource = "SELECT TagName , SPAGNOLO AS Valore FROM TestiMultilingua"
            Case 5
                AdoTesti.RecordSource = "SELECT TagName , TEDESCO AS Valore FROM TestiMultilingua"
            Case 6
                AdoTesti.RecordSource = "SELECT TagName , LinguaSpeciale AS Valore FROM TestiMultilingua"
        End Select
        AdoTesti.Refresh
        
        ' chiusura dei 2 database
        AdoTesti.Recordset.ActiveConnection = Nothing
        .Recordset.ActiveConnection = Nothing
    End With
    
End Sub

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   If Index <> frmKernel.PaginaCorrente Then
    '  Set AdoStoricoAllarme.Recordset.ActiveConnection = Nothing
   End If
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub Check1_Click()
   If One Then DB465.Bit(34, 1) = Not DB465.Bit(34, 1)
  ' FrameLav.Visible = CBool(Check1.Value)
End Sub

Private Sub ComdPreset_Click()
    DB465.Bit(34, 0) = True
End Sub

Private Sub Command1_Click()
  Dim x As Printer
  
  PrinterInstall = False
  For Each x In Printers
      If (InStr(x.DeviceName, Combo1.List(Combo1.ListIndex)) <> 0) And Combo1.List(Combo1.ListIndex) <> "Nothing" Then
         'If X.Orientation = vbPRORPortrait Then
         ' Imposta la stampante come predefinita di sistema.
         Set Printer = x
         PrinterInstall = True
         ' Interrompe la ricerca di una stampante.
         Exit For
      End If
   Next
End Sub

Private Sub CommandChiudi_Click()
    FrameGrid.Visible = False
End Sub
Private Sub Form_Activate()
   Dim i As Integer

   If One = False Then
       SSTab1.Visible = False

       Barra21.Selezionato = 2
       NumPagAlarmLog = 1
       '==================================
       ' imposta la lingua
       OptionLingua(GetNumber(TagLingua)).value = True
       On Error Resume Next
       For i = 0 To 5
          Shape4(i).BackColor = vbWhite
       Next
       On Error GoTo 0
       Shape4(GetNumber(TagLingua)).BackColor = vbYellow
       '=========================================
       One = True
       TimerLocale.Interval = 500
       TimerLocale.Enabled = True
   End If

End Sub
Private Sub Form_Load()
 '  Dim cn As ADODB.Connection
   Dim rs As ADODB.Recordset
   Dim StringaSql As String
   Dim i As Integer
   
   LeggiTabelle
   Frame4.ZOrder
   
   '============== aggiornamento lista stampanti =======
   
   Combo1.Clear
   On Error Resume Next
   Combo1.AddItem "Nothing"
   If UBound(PrintersList) > 0 Then
      For i = 1 To UBound(PrintersList)
         Combo1.AddItem PrintersList(i)
      Next
      Combo1.Text = Combo1.List(1)
   Else
      Combo1.Text = Combo1.List(0)
   End If
       
   '=====================================================
 '  Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
 '  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
   With rs
       .CursorLocation = adUseClient
       StringaSql = "SELECT * FROM Turni"
       .Open StringaSql, Connessione, adOpenForwardOnly, adLockReadOnly, adCmdText
       Set .ActiveConnection = Nothing
       'cn.Close
       For i = 1 To 3
          LblTurnoInizio(i) = .Fields("TurnoInizio")
          LblTurnoFine(i) = .Fields("TurnoFine")
          LblAliasTurno(i) = .Fields("TurnoAlias")
          .MoveNext
       Next
       .Close
   End With
   Set rs = Nothing
   'Set cn = Nothing

End Sub
Public Function Text(TagName As String) As String
    On Error Resume Next
    Text = "***"
    If TagName <> "" Then
        AdoTesti.Recordset.MoveFirst
        AdoTesti.Recordset.Find ("TagName = '" & TagName & "'")
        If AdoTesti.Recordset.EOF = False Then
            Text = AdoTesti.Recordset.Fields("Valore")
        Else
            MsgBox "Manca testo """ & TagName & """"
        End If
    End If
End Function
Public Function GetBit(TagName As String) As Boolean
    On Error Resume Next
    GetBit = False
    If TagName <> "" Then
        AdoNumeri.Recordset.MoveFirst
        AdoNumeri.Recordset.Find ("ID = '" & TagName & "'")
        If AdoNumeri.Recordset.EOF = False Then
            If AdoNumeri.Recordset.Fields("Valore") <> 0# Then
                GetBit = True
            End If
        Else
            MsgBox "Manca parametro """ & TagName & """"
        End If
    End If
End Function
Public Sub SetBit(TagName As String, valore As Boolean)
    If TagName <> "" Then
        AdoNumeri.Refresh
        AdoNumeri.Recordset.MoveFirst
        AdoNumeri.Recordset.Find ("ID = '" & TagName & "'")
        If AdoNumeri.Recordset.EOF = False Then
            If valore Then
                AdoNumeri.Recordset.Fields("Valore") = 1
            Else
                AdoNumeri.Recordset.Fields("Valore") = 0
            End If
            AdoNumeri.Recordset.Update
        Else
            MsgBox "Manca parametro """ & TagName & """"
        End If
        AdoNumeri.Recordset.ActiveConnection = Nothing
    End If
End Sub
Public Function GetNumber(TagName As String) As Double
    On Error Resume Next
    GetNumber = 0#
    If TagName <> "" Then
        AdoNumeri.Recordset.MoveFirst
        AdoNumeri.Recordset.Find ("ID = '" & TagName & "'")
        If AdoNumeri.Recordset.EOF = False Then
            GetNumber = AdoNumeri.Recordset.Fields("Valore")
        Else
            MsgBox "Manca parametro """ & TagName & """"
        End If
    End If
End Function

Public Sub SetNumber(TagName As String, valore As Double)
    If TagName <> "" Then
        AdoNumeri.Refresh
        AdoNumeri.Recordset.MoveFirst
        AdoNumeri.Recordset.Find ("ID = '" & TagName & "'")
        If AdoNumeri.Recordset.EOF = False Then
            AdoNumeri.Recordset.Fields("Valore") = valore
            AdoNumeri.Recordset.Update
        Else
            MsgBox "Manca parametro """ & TagName & """"
        End If
        AdoNumeri.Recordset.ActiveConnection = Nothing
    End If
    
    If TagName = TagLingua Then
        LeggiTabelle
    End If
End Sub

Private Sub GridParametri_Click()
  LblHelp.caption = HelpParam(GridParametri.Row)
End Sub

Private Sub LblAliasTurno_Click(Index As Integer)
    On Error Resume Next
    TOUCHKeyBoard.Dati = LblAliasTurno(Index).caption
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
            LblAliasTurno(Index).caption = TOUCHKeyBoard.Dati
            DownloadDatiTurno
    End If
   ' Update
End Sub

Private Sub LblTurnoFine_Click(Index As Integer)
   On Error Resume Next
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMax = 23
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.Dati = LblTurnoFine(Index).caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
            LblTurnoFine(Index).caption = TOUCHNumericPad.Dati
            DownloadDatiTurno
    End If
   ' Update
End Sub
Private Sub LblTurnoInizio_Click(Index As Integer)
    On Error Resume Next
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMax = 23
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.Dati = LblTurnoInizio(Index).caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            LblTurnoInizio(Index).caption = TOUCHNumericPad.Dati
            DownloadDatiTurno
        On Error GoTo 0
    End If
    'Update
End Sub

Private Sub Maxsoft_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = 20
'    TOUCHNumericPad.ValoreMax = 100
'    TOUCHNumericPad.Dati = DB460.Word(22)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati
'        If temp > 100# Then temp = 100#
'        If temp < 20 Then temp = 20#
'        DB460.Word(22) = temp
'    End If
End Sub

Private Sub Minsoft_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = 20
'    TOUCHNumericPad.ValoreMax = 100
'    TOUCHNumericPad.Dati = DB460.Word(22)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati
'        If temp > 100# Then temp = 100#
'        If temp < 20 Then temp = 20#
'        DB460.Word(22) = temp
'    End If
End Sub

Private Sub Offset_Click()
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = -15000
    TOUCHNumericPad.ValoreMax = 15000
    TOUCHNumericPad.Dati = DB465.Real(48)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        DB465.Real(48) = TOUCHNumericPad.Dati
        Offset = DB465.Real(48)
    End If
End Sub

Private Sub OptionLingua_Click(Index As Integer)
    Shape4(GetNumber(TagLingua)).BackColor = vbWhite
    SetNumber TagLingua, CDbl(Index)
    Shape4(Index).BackColor = vbYellow
    frmKernel.CaricaPagine True
    Me.Hide
    Me.Show
End Sub
Private Sub LabelOra_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1
    TOUCHNumericPad.ValoreMax = 23
    TOUCHNumericPad.Dati = LabelOra.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Time = TOUCHNumericPad.Dati & ":" & LabelMinuti.caption
        On Error GoTo 0
    End If
    'Update
End Sub
Private Sub LabelMinuti_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 59
    TOUCHNumericPad.Dati = LabelMinuti.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Time = LabelOra.caption & ":" & TOUCHNumericPad.Dati
        On Error GoTo 0
    End If
   ' Update
End Sub
Private Sub LabelGiorno_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMax = 31
    TOUCHNumericPad.ValoreMin = 1
    TOUCHNumericPad.Dati = LabelGiorno.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Date = TOUCHNumericPad.Dati & "/" & LabelMese.caption & "/" & LabelAnno.caption
        On Error GoTo 0
    End If
    'Update
End Sub
Private Sub LabelMese_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1
    TOUCHNumericPad.ValoreMax = 12
    TOUCHNumericPad.Dati = LabelMese.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Date = LabelGiorno.caption & "/" & TOUCHNumericPad.Dati & "/" & LabelAnno.caption
        On Error GoTo 0
    End If
    'Update
End Sub
Private Sub LabelAnno_Click()
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 1900
    TOUCHNumericPad.ValoreMax = 3000
    TOUCHNumericPad.Dati = LabelAnno.caption
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        On Error Resume Next
            Date = LabelGiorno.caption & "/" & LabelMese.caption & "/" & TOUCHNumericPad.Dati
        On Error GoTo 0
    End If
    'Update
End Sub
Public Sub Update()
     ' aggiorna lo stato del pulsante comunicazione
    On Error Resume Next
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    LabelOra.caption = Format(Now, "hh")
    LabelMinuti.caption = Mid(Format(Now, "hh:mm"), 4, 2)
    LabelGiorno.caption = Format(Now, "dd")
    LabelMese.caption = Format(Now, "mm")
    LabelAnno.caption = Format(Now, "yyyy")
    If UBound(PrintersList) > 0 Then
       Label6.caption = Printer.DeviceName
    Else
       Label6.caption = Combo1.List(0)
    End If
    actual = DB418.Word(54)
    Offset = DB465.Real(48)
    Preset = DB465.Real(44)
    Check1.value = Abs(Int(DB465.Bit(34, 1)))
End Sub
Sub ScriviParametriSuPlc()
    On Error Resume Next
    DB422.Word(34) = Param.GetNumber("Par223_OffSetRegg")
    DB402.Word(10) = Param.GetNumber("DB402_014_BandaPesaAZero")
    DB480.Word(90) = Param.GetNumber("DB480_090_ParF_GiriAnelloTesta")
    DB480.Word(92) = Param.GetNumber("DB480_092_ParF_GiriAnelloCoda")
    DB480.Word(94) = Param.GetNumber("DB480_094_ParTL_CoeffConteggio1")
    DB480.Word(96) = Param.GetNumber("DB480_096_ParTL_CoeffConteggio2")
    DB480.Word(98) = Param.GetNumber("DB480_098_ParTL_CoeffCentrPacco")
    DB480.Word(100) = Param.GetNumber("DB480_100_ParTL_QuotaRallenta")
    DB480.Word(102) = Param.GetNumber("DB480_102_ParTL_OffsetReggiat")
    DB480.Word(104) = Param.GetNumber("DB480_104_ParTL_QuotaInterF_R")
    DB480.Word(106) = Param.GetNumber("DB480_106_ParF_QuotaFascTesta")
    DB480.Word(108) = Param.GetNumber("DB480_108_ParF_QuotaFascCoda")
    On Error GoTo 0
End Sub

Private Sub CommandModifica_Click()
    LblHelp.caption = HelpParam(GridParametri.Row)
    GridParametri.Col = 3    ' ID parametro è sulla colonna 0
    TOUCHNumericPad.Decimali = 2
    TOUCHNumericPad.ValoreMax = 999999999
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.Dati = GetNumber(GridParametri.Text)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        SetNumber GridParametri.Text, TOUCHNumericPad.Dati
        LeggiTabelle
        ScriviParametriSuPlc
       ' GridParametri.Refresh
       AggiornaTabParametri
    End If
End Sub

Private Sub CommandUp_Click()
    Dim PosizioneAttuale As Integer
    PosizioneAttuale = GridParametri.TopRow
    PosizioneAttuale = PosizioneAttuale - 10
    If PosizioneAttuale < 0 Then PosizioneAttuale = 0
    If GridParametri.Rows > 0 Then GridParametri.TopRow = PosizioneAttuale
End Sub
Private Sub CommandDown_Click()
    Dim PosizioneAttuale As Integer
    PosizioneAttuale = GridParametri.TopRow
    PosizioneAttuale = PosizioneAttuale + 10
    If PosizioneAttuale >= GridParametri.Rows Then PosizioneAttuale = GridParametri.Rows - 1
    If GridParametri.Rows > 0 Then GridParametri.TopRow = PosizioneAttuale
End Sub

Private Sub Preset_Click()
   ' Dim temp As Double
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 15000
    TOUCHNumericPad.Dati = DB465.Real(44)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        DB465.Real(44) = TOUCHNumericPad.Dati
        Preset = DB465.Real(44)
    End If
End Sub

Private Sub TimerLocale_Timer()
   Me.Update
End Sub
Private Sub AggiornaTabParametri()
  Dim cn As New ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim LinguaPar As String
    
    GridParametri.Cols = 4
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Parameters.mdb;Persist Security Info=False"
    FrameGrid.Visible = True
    FrameGrid.ZOrder 0
    DoEvents
    With rs
        .Open "select * from Parametri order by Gruppo DESC,ID", cn, , adLockReadOnly, adCmdText
        ReDim HelpParam(.RecordCount + 1)
        IndiceRiga = 0
        GridParametri.ColWidth(0) = 5000
        GridParametri.ColWidth(1) = 1000
        GridParametri.ColWidth(2) = 5000
        GridParametri.ColWidth(3) = 5000
        '========================================================================
        ' selezione lingua parametri
        Select Case GetNumber(TagLingua)
        Case 1
                  LinguaPar = "ITA"
        Case 2
                  LinguaPar = "ING"
        Case 3
                  LinguaPar = "FRA"
        Case 4
                  LinguaPar = "SPA"
        Case 5
                  LinguaPar = "TED"
        Case 6
                  LinguaPar = "SPC"
        End Select
        '========================================================================
        While .EOF = False
            ReDim Preserve HelpParam(IndiceRiga + 1)
            If GridParametri.Rows < (IndiceRiga + 1) Then GridParametri.AddItem ""
            GridParametri.Row = IndiceRiga
            GridParametri.RowHeight(IndiceRiga) = 450
            On Error Resume Next
                HelpParam(IndiceRiga) = .Fields("Help")
                GridParametri.Col = 0
                GridParametri.CellBackColor = Gruppo(.Fields("Gruppo"))
                GridParametri.CellAlignment = flexAlignLeftCenter
                GridParametri.Text = .Fields(LinguaPar)
                
                GridParametri.Col = 1
                GridParametri.CellBackColor = Gruppo(.Fields("Gruppo"))
                GridParametri.CellAlignment = flexAlignCenterCenter
                GridParametri.Text = .Fields("Valore")
                
                GridParametri.Col = 2
                GridParametri.CellBackColor = Gruppo(.Fields("Gruppo"))
                GridParametri.CellAlignment = flexAlignLeftCenter
                GridParametri.Text = .Fields("Selezioni")
                
                GridParametri.Col = 3
                GridParametri.CellBackColor = Gruppo(.Fields("Gruppo"))
                GridParametri.CellAlignment = flexAlignLeftCenter
                GridParametri.Text = .Fields("ID")
            On Error GoTo 0
            
            IndiceRiga = IndiceRiga + 1
            .MoveNext
        Wend
    End With
    GridParametri.Rows = IndiceRiga
    Set rs = Nothing
    Set cn = Nothing
End Sub
Private Sub VerticalMenu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
 On Error Resume Next
    Select Case MenuNumber
        Case 1
            Select Case MenuItem
                Case 1
                      Shell "..\target\DatabaseEditor.exe", vbNormalFocus
                Case 2
                      Shell "..\target\HelpEd.exe", vbNormalFocus
                Case 3
                     On Error GoTo InstallSoftNet
                     Shell "C:\WINNT\system32\s7epatsx.exe -App=Simatic", vbNormalFocus
                     On Error GoTo 0
                     Exit Sub
InstallSoftNet: MsgBox "Installare Pb soft net s7 V5.3", vbCritical, "Installazioni"
                     Exit Sub
                Case 4
                      On Error GoTo InstallSoftNetCom
                       Shell "C:\SIEMENS\SIMATIC.NET\coms7.nt\COMLS7.EXE -leng", vbNormalFocus
                       On Error GoTo 0
                       Exit Sub
InstallSoftNetCom: MsgBox "Installare Pb soft net s7 V5.3", vbCritical, "Installazioni"
                       Exit Sub
                Case 5
                        On Error Resume Next
                        Shell "C:\SIEMENS\Step7\s7bin\S7tgtopx.exe"
                        Exit Sub
                Case 6
                Case 7
            End Select
        Case 2
            Select Case MenuItem
                Case 1
                        Dim Risposta
                        Risposta = MsgBox("Sure ?", vbYesNo, "DP6 - Download database data in plc")
                        If Risposta = vbYes Then
                           frmMovingData.FromPLC = False
                           frmMovingData.Show
                           frmKernel.CaricaDatiDBSimulazioneInPlc
                           frmMovingData.Hide
                        End If
                Case 2
                           frmMovingData.FromPLC = True
                           frmMovingData.Show
                            frmKernel.SalvaDatiPLCInDBSimulazione
                           frmMovingData.Hide
                Case 3
                Case 4
                Case 5
            End Select
        Case 3
            Select Case MenuItem
                Case 1
                       On Error GoTo ErrorePercorso1
                        frmHelp.NomeFile = "TIP.HTM"
                        frmHelp.Contesto = "Impostazioni Driver siemens"
                        frmHelp.Show
                        FrameGrid.Visible = True
                        FrameGrid.ZOrder 0
                        frmHelp.Top = 0
                        frmHelp.Left = 7500
                        Exit Sub
ErrorePercorso1:
                        MsgBox "Percorso file guida errato", vbExclamation
                        Exit Sub
                Case 2
                        On Error GoTo ErrorePercorso2
                        frmHelp.NomeFile = "TIP.HTM"
                        frmHelp.Contesto = "Com LOG"
                        frmHelp.Show
                        FrameGrid.Visible = True
                        FrameGrid.ZOrder 0
                        frmHelp.Top = 0
                        frmHelp.Left = 7500
                        Exit Sub
ErrorePercorso2:
                        MsgBox "Percorso file guida errato", vbExclamation
                    Exit Sub
                Case 3
                Case 4
                Case 5
                Case 6
            End Select
        Case 4
            Select Case MenuItem
                Case 1
                    Shell "explorer.exe", vbNormalFocus
                Case 2
                    Shell "calc.exe", vbNormalFocus
                Case 3
                     Shell "notepad.exe", vbNormalFocus
                Case 4
                     Shell "mspaint.exe", vbNormalFocus
                Case 5
            End Select
       Case 5 ' menu DP6
        Select Case MenuItem
                Case 1
                    Shell "..\target\RSM.exe", vbNormalFocus
                Case 2
                    If Param.GetBit("Par211_DownloadPLCDataInDB") Then
                      frmMovingData.FromPLC = True
                      frmMovingData.Show
                      frmKernel.SalvaDatiPLCInDBSimulazione
                      frmMovingData.Hide
                    End If
                    frmKernel.ChiusuraProgramma
                    End
                Case 3
                Case 4
                Case 5
            End Select
    End Select
End Sub
Function Gruppo(ByVal n As Integer) As Variant
Select Case n
  Case 1
         Gruppo = &H80FFFF
  Case 2
         Gruppo = &HFF8080
  Case 3
         Gruppo = &H80FF&
  Case 4
         Gruppo = &H80FF80
End Select
End Function
Sub ScritteMultilingua()
   'Command1.Caption = Text("AnalisiAllarmi")
   'Command3.Caption = Text("DefTurni")
   XPButton1(5).TxtText = Text("ParDp6")
   Label1 = Text("Turno") & " 1"
   Label2 = Text("Turno") & " 2"
   Label3 = Text("Turno") & " 3"
   Label8 = Text("DefTurni")
   Label4 = Text("Lingua")
   Label5 = Text("ModDataora")
  ' XPButton1(4).TxtText = Text("CancLogPacco")
   lblbar(5) = "Service" 'Param.Text("Orders page")
   lblbar(1) = Param.Text("Pagina")
  ' XPButton1(6).TxtText = Param.Text("CancLogAllarmi")
   'lblbar(3) = Param.Text("Ricette")
   'lblbar(0) = Param.Text("ORDER")
   '================================
   LabelData.caption = Param.Text("Data")
   LabelOre.caption = Param.Text("Ora")
   CommandModifica.caption = Param.Text("MODIFICA")
   CommandChiudi.caption = Param.Text("Chiudi")
End Sub
Sub DownloadDatiTurno()
  ' Dim cn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim StringaSql As String
   Dim i As Integer
   
        '  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
          With rs
              StringaSql = "SELECT * FROM Turni"
              .Open StringaSql, Connessione, adOpenForwardOnly, adLockOptimistic, adCmdText
              For i = 1 To 3
                   .Fields("TurnoInizio") = LblTurnoInizio(i)
                   DatiTurno(i, 1) = LblTurnoInizio(i)
                   .Fields("TurnoFine") = LblTurnoFine(i)
                   DatiTurno(i, 2) = LblTurnoFine(i)
                   .Fields("TurnoAlias") = LblAliasTurno(i)
                   DatiTurno(i, 3) = LblAliasTurno(i)
                   .MoveNext
              Next
              .Close
              Set .ActiveConnection = Nothing
          End With
          Set rs = Nothing
  '        Set cn = Nothing
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
                .NomeFile = "service_pagina.htm"
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .webPreview.Move 0, 0, frmHelp.Width, frmHelp.Height
               .Show
           End With
    Case 4
          TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
          TechPasswordForm.Show vbModal
          If TechPasswordForm.LoginSucceeded = False Then Exit Sub
          Unload TechPasswordForm
    
          SSTab1.Visible = True
          SSTab1.ZOrder
    Case 5
          Dim IndiceRiga As Integer
         
          TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par110_Password"))
          TechPasswordForm.Show vbModal
          If TechPasswordForm.LoginSucceeded = False Then Exit Sub
          Unload TechPasswordForm
          LeggiTabelle
          AggiornaTabParametri
    Case 6
'          On Error GoTo ErrorePercorso
'          Unload frmHelp
'          Set frmHelp = Nothing
'          frmHelp.Errori = False
'          frmHelp.NomeFile = "TIP.HTM"
'          frmHelp.Contesto = "DP6 : CP_L2_1 COM LOG"
'          frmHelp.Top = 0
'          frmHelp.Left = 7500
'          frmHelp.Show vbModal
    Case 7
         On Error GoTo ErrorePercorso
           Unload frmHelp
           Set frmHelp = Nothing
           With frmHelp
                .Errori = True
                .NomeFile = "2202a.pdf"
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

Private Sub XPButton2_Click()
      frmStampa.Show
End Sub
