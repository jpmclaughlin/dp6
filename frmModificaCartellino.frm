VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmModificaCartellino 
   BorderStyle     =   0  'None
   Caption         =   "&H00E0E0E0&"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionLingua 
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2760
      Width           =   2355
   End
   Begin VB.OptionButton OptionLingua 
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3570
      Width           =   2355
   End
   Begin VB.OptionButton OptionLingua 
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4350
      Width           =   2355
   End
   Begin VB.OptionButton OptionLingua 
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5130
      Width           =   2355
   End
   Begin VB.OptionButton OptionLingua 
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
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5910
      Value           =   -1  'True
      Width           =   2355
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9795
      Left            =   360
      TabIndex        =   4
      Top             =   1620
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   17277
      _Version        =   393216
      Tab             =   1
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
      TabPicture(0)   =   "frmModificaCartellino.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "campo(3)"
      Tab(0).Control(1)=   "campo(4)"
      Tab(0).Control(2)=   "campo(5)"
      Tab(0).Control(3)=   "campo(6)"
      Tab(0).Control(4)=   "campo(7)"
      Tab(0).Control(5)=   "campo(8)"
      Tab(0).Control(6)=   "campo(9)"
      Tab(0).Control(7)=   "campo(10)"
      Tab(0).Control(8)=   "campo(2)"
      Tab(0).Control(9)=   "campo(1)"
      Tab(0).Control(10)=   "Shape2(12)"
      Tab(0).Control(11)=   "Line1(9)"
      Tab(0).Control(12)=   "Line1(8)"
      Tab(0).Control(13)=   "Line1(7)"
      Tab(0).Control(14)=   "Line1(6)"
      Tab(0).Control(15)=   "Line1(5)"
      Tab(0).Control(16)=   "Line1(4)"
      Tab(0).Control(17)=   "Line1(3)"
      Tab(0).Control(18)=   "Line1(2)"
      Tab(0).Control(19)=   "Line1(1)"
      Tab(0).Control(20)=   "Line1(0)"
      Tab(0).Control(21)=   "Text1(3)"
      Tab(0).Control(22)=   "Text1(4)"
      Tab(0).Control(23)=   "Text1(5)"
      Tab(0).Control(24)=   "Text1(6)"
      Tab(0).Control(25)=   "Text1(7)"
      Tab(0).Control(26)=   "Text1(8)"
      Tab(0).Control(27)=   "Text1(9)"
      Tab(0).Control(28)=   "Text1(10)"
      Tab(0).Control(29)=   "Text1(2)"
      Tab(0).Control(30)=   "Text1(1)"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Automatici"
      TabPicture(1)   =   "frmModificaCartellino.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape2(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "campo(11)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "campo(12)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "campo(20)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "campo(19)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "campo(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "campo(17)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "campo(16)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "campo(15)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "campo(14)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "campo(13)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line1(10)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Line1(11)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Line1(12)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Line1(13)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Line1(14)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Line1(15)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Line1(16)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Line1(17)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Line1(18)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Line1(19)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text1(11)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text1(12)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text1(20)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text1(19)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text1(18)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Text1(17)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Text1(16)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text1(15)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text1(14)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text1(13)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "Anteprima"
      TabPicture(2)   =   "frmModificaCartellino.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5805
         Left            =   -71610
         ScaleHeight     =   5805
         ScaleWidth      =   5055
         TabIndex        =   50
         Top             =   690
         Width           =   5055
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
         Left            =   5310
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   2670
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
         Left            =   5310
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   3360
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
         Left            =   5310
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   4050
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
         Index           =   16
         Left            =   5310
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   4740
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
         Index           =   17
         Left            =   5310
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   5430
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
         Index           =   18
         Left            =   5310
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   6120
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
         Index           =   19
         Left            =   5310
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   6810
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
         Index           =   20
         Left            =   5340
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   7500
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
         Index           =   12
         Left            =   5310
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1980
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
         Index           =   11
         Left            =   5310
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1290
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
         Left            =   -69720
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1260
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
         Left            =   -69720
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1950
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
         Left            =   -69720
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   7470
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
         Left            =   -69720
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   6780
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
         Left            =   -69720
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   6090
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
         Left            =   -69720
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   5400
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
         Left            =   -69720
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4710
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
         Left            =   -69720
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4020
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
         Left            =   -69720
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3330
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
         Left            =   -69720
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2640
         Width           =   6375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   -74640
         X2              =   -69240
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   1
         X1              =   -74640
         X2              =   -69240
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   2
         X1              =   -74580
         X2              =   -69180
         Y1              =   6690
         Y2              =   6690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   3
         X1              =   -74580
         X2              =   -69180
         Y1              =   6030
         Y2              =   6030
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   4
         X1              =   -74580
         X2              =   -69180
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   5
         X1              =   -74580
         X2              =   -69180
         Y1              =   4650
         Y2              =   4650
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   6
         X1              =   -74640
         X2              =   -69240
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   7
         X1              =   -74640
         X2              =   -69240
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   8
         X1              =   -74580
         X2              =   -69180
         Y1              =   7410
         Y2              =   7410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   9
         X1              =   -74580
         X2              =   -69180
         Y1              =   8100
         Y2              =   8100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   19
         X1              =   360
         X2              =   5760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   18
         X1              =   450
         X2              =   5850
         Y1              =   8130
         Y2              =   8130
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   17
         X1              =   450
         X2              =   5850
         Y1              =   7470
         Y2              =   7470
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   16
         X1              =   390
         X2              =   5790
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   15
         X1              =   390
         X2              =   5790
         Y1              =   3990
         Y2              =   3990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   14
         X1              =   450
         X2              =   5850
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   13
         X1              =   450
         X2              =   5850
         Y1              =   5370
         Y2              =   5370
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   12
         X1              =   450
         X2              =   5850
         Y1              =   6090
         Y2              =   6090
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   11
         X1              =   450
         X2              =   5850
         Y1              =   6750
         Y2              =   6750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   10
         X1              =   390
         X2              =   5790
         Y1              =   2610
         Y2              =   2610
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Descrizione"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   270
         TabIndex        =   49
         Top             =   2670
         Width           =   4695
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensione tubo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   14
         Left            =   270
         TabIndex        =   48
         Top             =   3360
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lunghezza tubo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   15
         Left            =   270
         TabIndex        =   47
         Top             =   4080
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero pacco"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   16
         Left            =   270
         TabIndex        =   46
         Top             =   4740
         Width           =   4785
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero tubi"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   17
         Left            =   270
         TabIndex        =   45
         Top             =   5430
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Peso pacco"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   18
         Left            =   270
         TabIndex        =   44
         Top             =   6120
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Spessore tubo"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   19
         Left            =   270
         TabIndex        =   43
         Top             =   6780
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   20
         Left            =   270
         TabIndex        =   42
         Top             =   7500
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ora"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   270
         TabIndex        =   41
         Top             =   1980
         Width           =   4635
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   270
         TabIndex        =   40
         Top             =   1320
         Width           =   4755
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   7635
         Index           =   2
         Left            =   5010
         Shape           =   4  'Rounded Rectangle
         Top             =   810
         Width           =   7065
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   7635
         Index           =   12
         Left            =   -70020
         Shape           =   4  'Rounded Rectangle
         Top             =   780
         Width           =   7065
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -74760
         TabIndex        =   24
         Top             =   1290
         Width           =   4695
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -74760
         TabIndex        =   23
         Top             =   1950
         Width           =   4695
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   -74760
         TabIndex        =   22
         Top             =   7470
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   -74760
         TabIndex        =   21
         Top             =   6750
         Width           =   4725
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   -74760
         TabIndex        =   20
         Top             =   6090
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   -74760
         TabIndex        =   19
         Top             =   5400
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   -74760
         TabIndex        =   18
         Top             =   4710
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   -74760
         TabIndex        =   17
         Top             =   4050
         Width           =   4755
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   -74760
         TabIndex        =   16
         Top             =   3330
         Width           =   4725
      End
      Begin VB.Label campo 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "campo 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   -74760
         TabIndex        =   15
         Top             =   2640
         Width           =   4755
      End
   End
   Begin VB.CommandButton CmdAnnulla 
      Caption         =   "Annulla"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12900
      TabIndex        =   3
      Top             =   10440
      Width           =   2085
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12900
      TabIndex        =   2
      Top             =   9540
      Width           =   2085
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   8985
      Index           =   1
      Left            =   12660
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   2565
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EE3959&
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket modify page"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   825
      Left            =   4500
      TabIndex        =   0
      Top             =   30
      Width           =   7545
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00D8B1AF&
      BackStyle       =   1  'Opaque
      Height          =   11820
      Index           =   0
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   870
      Width           =   15315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00EE3959&
      BackStyle       =   0  'Transparent
      Caption         =   "Ticket modify page   "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   885
      Index           =   2
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   8985
   End
End
Attribute VB_Name = "frmModificaCartellino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Private LinguaSelezionata As Byte
Private OrdineTemp As OrderClass
Private Anteprima As PictureBox

''==============================================================================
'' FUNZIONE AGGIORNAMENTO pagina
''==============================================================================
'
'Sub Update()
'    Dim i As Integer
'
'    ' crea la variabile ordine corrente
'    Set OrdineTemp = New OrderClass
'    'imposta il valore della variabile ordine corrente al valore corrente
'    OrdineTemp.IDOrdine = 1
'    If frmKernel.CodOrdineCorrente.CodStoccaggio <> 0 Then
'       OrdineTemp.IDOrdine = frmKernel.CodOrdineCorrente.CodStoccaggio
'    End If
'    'carica dal database i dati riguardanti l'ordine corrente
'    OrdersForm.LoadOrderData OrdineTemp
'    'FormStampa.AggiornamentoCampi
'
'    ' update della pagina con i valori caricati
'    For i = 1 To 20
'          If FormStampa.CampoVariabile(i).Visible = True Then
'             Text1(i).Visible = True
'             campo(i).Visible = True
'             If i < 11 Then
'                Text1(i).Text = OrdineTemp.CampoManuale(i)
'                campo(i).Caption = FormStampa.CampoFisso(i)
'            ' Else
'                'Text1(i).Text = FormStampa.CampoFisso(i).Caption
'             End If
'          Else
'             Text1(i).Visible = False
'             campo(i).Visible = False
'          End If
'    Next
'
'    Text1(11).Text = FormStampa.CampoVariabile(11).Caption
'    Text1(12).Text = FormStampa.CampoVariabile(12).Caption
'    Text1(13).Text = FormStampa.CampoVariabile(13).Caption
'    Text1(14).Text = FormStampa.CampoVariabile(14).Caption
'    Text1(15).Text = FormStampa.CampoVariabile(15).Caption
'    Text1(16).Text = FormStampa.CampoVariabile(16).Caption
'    Text1(17).Text = FormStampa.CampoVariabile(17).Caption
'    Text1(18).Text = FormStampa.CampoVariabile(18).Caption
'    Text1(19).Text = FormStampa.CampoVariabile(19).Caption
'
'    OptionLingua(FormStampa.OrderTemp.LinguaCartellino).Value = True
'End Sub

'Private Sub CmdAnnulla_Click()
'   Set OrdineTemp = Nothing
'   Unload Me
'End Sub

'Private Sub cmdOK_Click()
'Dim i As Integer
'
'   OrdineTemp.LinguaCartellino = LinguaSelezionata
'   For i = 1 To 10
'      OrdineTemp.CampoManuale(i) = Text1(i).Text
'      Cartellino.CampoManuale(i) = Text1(i).Text
'   Next
'
'   OrdersForm.SaveOrderData OrdineTemp
'   Set OrdineTemp = Nothing
'   Unload Me
'End Sub

'Private Sub Command1_Click()
'   frmAnteprima.Show vbModal
'End Sub
'
'Private Sub Form_Activate()
'   DrawAnteprima
'End Sub
'
'Private Sub Form_Load()
'
'    FormStampa.AggiornamentoCampi
'    Label2.Caption = Param.Text("Ticket modify page")
'    Label1(2).Caption = Param.Text("Ticket modify page")
'    Me.Update
'    WindowState = vbMaximized
'    SSTab1.Tab = 0
'End Sub

'Private Sub OptionLingua_Click(Index As Integer)
'Dim i As Integer
'
'    Select Case Index
'        Case 1
'            FormStampa.CampoLingua = "ITALIANO"
'        Case 2
'            FormStampa.CampoLingua = "INGLESE"
'        Case 3
'            FormStampa.CampoLingua = "FRANCESE"
'        Case 4
'            FormStampa.CampoLingua = "SPAGNOLO"
'        Case 5
'            FormStampa.CampoLingua = "TEDESCO"
'        Case Else
'            FormStampa.CampoLingua = "ITALIANO"
'     End Select
'
'     LinguaSelezionata = Index
'     FormStampa.CaricaCampiFissi
'     ' update della pagina con i valori caricati
'     For i = 1 To 20
'        'Text1(i).Text = OrdineTemp.CampoManuale(i)
'        campo(i).Caption = FormStampa.CampoFisso(i)
'        If FormStampa.CampoFisso(i).Visible = True Then
'           Text1(i).Visible = True
'           campo(i).Visible = True
'        Else
'           Text1(i).Visible = False
'           campo(i).Visible = False
'        End If
'     Next
'     DrawAnteprima
'End Sub
'
'
'
'Private Sub SSTab1_Click(PreviousTab As Integer)
'   DrawAnteprima
'End Sub
'

'Private Sub Text1_Click(Index As Integer)
'    TOUCHKeyBoard.TextModifica.PasswordChar = ""
'    TOUCHKeyBoard.Dati = Text1(Index).Text
'    TOUCHKeyBoard.Show vbModal
'    If TOUCHKeyBoard.DatiConfermati Then
'        Text1(Index).Text = TOUCHKeyBoard.Dati
'        FormStampa.CampoVariabile(Index).Caption = TOUCHKeyBoard.Dati
'    End If
'    TOUCHKeyBoard.TextModifica.PasswordChar = ""
'End Sub

'Private Sub DrawAnteprima()
'    Dim i As Integer
'
'    Set Anteprima = Picture1
'
'    Picture1.Cls
'    With Anteprima
'        '.left = (SSTab1.Width / 2) - (FormStampa.ImgSfondo.Width / 2)
'        .Width = FormStampa.ImgSfondo.Width
'        .Height = FormStampa.Height
'        .Picture = FormStampa.ImgSfondo.Picture
'        For i = 1 To 20
'                If FormStampa.CampoFisso(i).Visible = True Then
'                    .FontName = FormStampa.CampoFisso(i).FontName
'                    .FontSize = FormStampa.CampoFisso(i).FontSize
'                    .FontBold = FormStampa.CampoFisso(i).FontBold
'                    .CurrentX = FormStampa.CampoFisso(i).left
'                    .CurrentY = FormStampa.CampoFisso(i).tOp
'                    Anteprima.Print FormStampa.CampoFisso(i).Caption
'
'                    .FontName = FormStampa.CampoVariabile(i).FontName
'                    .FontSize = FormStampa.CampoVariabile(i).FontSize
'                    .FontBold = FormStampa.CampoVariabile(i).FontBold
'                    .CurrentX = FormStampa.CampoVariabile(i).left
'                    .CurrentY = FormStampa.CampoVariabile(i).tOp
'                    Anteprima.Print FormStampa.CampoVariabile(i).Caption
'                End If
'        Next
'        For i = 0 To 2
'              If FormStampa.Intestazione(i).Visible = True Then
'                    .FontName = FormStampa.Intestazione(i).FontName
'                    .FontSize = FormStampa.Intestazione(i).FontSize
'                    .FontBold = FormStampa.Intestazione(i).FontBold
'                    .CurrentX = FormStampa.Intestazione(i).left
'                    .CurrentY = FormStampa.Intestazione(i).tOp
'                    Anteprima.Print FormStampa.Intestazione(i).Caption
'            End If
'        Next
'    End With
'End Sub
'
Private Sub Form_Load()

End Sub

