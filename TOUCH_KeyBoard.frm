VERSION 5.00
Begin VB.Form TOUCHKeyBoard 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10140
   ClientLeft      =   315
   ClientTop       =   510
   ClientWidth     =   14610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   14610
   Begin VB.CommandButton KeyClear 
      BackColor       =   &H8000000B&
      Cancel          =   -1  'True
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8115
      Width           =   3375
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   311
      Left            =   13260
      TabIndex        =   53
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   310
      Left            =   12120
      TabIndex        =   52
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   211
      Left            =   12660
      TabIndex        =   51
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   210
      Left            =   11520
      TabIndex        =   50
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   112
      Left            =   13260
      TabIndex        =   49
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "["
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   111
      Left            =   12120
      TabIndex        =   48
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   12
      Left            =   12720
      TabIndex        =   47
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   11
      Left            =   11580
      TabIndex        =   46
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   401
      Left            =   3540
      TabIndex        =   45
      Top             =   6360
      Width           =   6795
   End
   Begin VB.CommandButton KeyArrowDX 
      BackColor       =   &H8000000A&
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10380
      TabIndex        =   44
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton KeyArrowSX 
      BackColor       =   &H8000000A&
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "System"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      TabIndex        =   43
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   """"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   309
      Left            =   10980
      TabIndex        =   42
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   308
      Left            =   9780
      TabIndex        =   41
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   301
      Left            =   1800
      TabIndex        =   40
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   302
      Left            =   2940
      TabIndex        =   39
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   303
      Left            =   4080
      TabIndex        =   38
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   304
      Left            =   5220
      TabIndex        =   37
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   305
      Left            =   6360
      TabIndex        =   36
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   306
      Left            =   7500
      TabIndex        =   35
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   307
      Left            =   8640
      TabIndex        =   34
      Top             =   5340
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   209
      Left            =   10380
      TabIndex        =   33
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   208
      Left            =   9240
      TabIndex        =   32
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   207
      Left            =   8100
      TabIndex        =   31
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   206
      Left            =   6960
      TabIndex        =   30
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   205
      Left            =   5820
      TabIndex        =   29
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   204
      Left            =   4680
      TabIndex        =   28
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   203
      Left            =   3540
      TabIndex        =   27
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   202
      Left            =   2400
      TabIndex        =   26
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   201
      Left            =   1260
      TabIndex        =   25
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   110
      Left            =   10980
      TabIndex        =   24
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   109
      Left            =   9840
      TabIndex        =   23
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   108
      Left            =   8700
      TabIndex        =   22
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   107
      Left            =   7560
      TabIndex        =   21
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   106
      Left            =   6420
      TabIndex        =   20
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   105
      Left            =   5280
      TabIndex        =   19
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   104
      Left            =   4140
      TabIndex        =   18
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "<< BS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Index           =   0
      Left            =   11790
      TabIndex        =   17
      Top             =   495
      Width           =   1935
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   10
      Left            =   10440
      TabIndex        =   16
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   7
      Left            =   7020
      TabIndex        =   15
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   8
      Left            =   8160
      TabIndex        =   14
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   9
      Left            =   9300
      TabIndex        =   13
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   4
      Left            =   3600
      TabIndex        =   12
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   5
      Left            =   4740
      TabIndex        =   11
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   6
      Left            =   5880
      TabIndex        =   10
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   3
      Left            =   2460
      TabIndex        =   7
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   101
      Left            =   720
      TabIndex        =   6
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   102
      Left            =   1860
      TabIndex        =   5
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton Key 
      BackColor       =   &H8000000A&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   103
      Left            =   3000
      TabIndex        =   4
      Top             =   3300
      Width           =   1100
   End
   Begin VB.CommandButton KeyCapsLock 
      BackColor       =   &H8000000B&
      Caption         =   "CAPS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12120
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   2235
   End
   Begin VB.TextBox TextModifica 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   300
      TabIndex        =   2
      Top             =   480
      Width           =   11355
   End
   Begin VB.CommandButton KeyEsc 
      BackColor       =   &H000000FF&
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5025
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8115
      Width           =   3375
   End
   Begin VB.CommandButton KeyEnter 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1335
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8115
      Width           =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   60
      X2              =   14565
      Y1              =   10095
      Y2              =   10110
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   75
      X2              =   14715
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   14580
      X2              =   14580
      Y1              =   30
      Y2              =   10170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      X1              =   15
      X2              =   15
      Y1              =   45
      Y2              =   10185
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   840
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "TOUCHKeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public DatiConfermati As Boolean
Public Dati As String

Private CapsLock As Boolean

Private Sub Form_Activate()
    DatiConfermati = False
    TextModifica.Text = Dati
    TextModifica.SetFocus
    SendKeys "{END}"
    CapsLock = True
    KeyCapsLock.BackColor = RGB(&HFF, &HFF, &H0) 'giallo
    Maiuscole
End Sub

Private Sub KeyEnter_Click()
    ' Restituisce la stringa togliendo gli eventuali spazi iniziali e finali
    Dati = Trim(TextModifica.Text)
    DatiConfermati = True
    Me.Hide
End Sub

Private Sub KeyEsc_Click()
    DatiConfermati = False
    Me.Hide
End Sub

Private Sub KeyClear_Click()
    TextModifica.SetFocus
    TextModifica.Text = ""
    End Sub

Private Sub Key_Click(Index As Integer)
    TextModifica.SetFocus
    Select Case Index
        Case 0
            SendKeys "{BS}"
        Case 1
            If CapsLock Then
                SendKeys "1"
            Else
                SendKeys "!"
            End If
        Case 2
            If CapsLock Then
                SendKeys "2"
            Else
                SendKeys "@"
            End If
        Case 3
            If CapsLock Then
                SendKeys "3"
            Else
                SendKeys "£"
            End If
        Case 4
            If CapsLock Then
                SendKeys "4"
            Else
                SendKeys "$"
            End If
        Case 5
            If CapsLock Then
                SendKeys "5"
            Else
                SendKeys "{%}"
            End If
        Case 6
            If CapsLock Then
                SendKeys "6"
            Else
                SendKeys "&"
            End If
        Case 7
            If CapsLock Then
                SendKeys "7"
            Else
                SendKeys "*"
            End If
        Case 8
            If CapsLock Then
                SendKeys "8"
            Else
                SendKeys "{(}"
            End If
        Case 9
            If CapsLock Then
                SendKeys "9"
            Else
                SendKeys "{)}"
            End If
        Case 10
            If CapsLock Then
                SendKeys "0"
            Else
                SendKeys "="
            End If
        Case 11
            If CapsLock Then
                SendKeys "-"
            Else
                SendKeys "{+}"
            End If
        Case 12
            If CapsLock Then
                SendKeys "/"
            Else
                SendKeys "\"
            End If
        
        Case 101
            If CapsLock Then
                SendKeys "Q"
            Else
                SendKeys "q"
            End If
        Case 102
            If CapsLock Then
                SendKeys "W"
            Else
                SendKeys "w"
            End If
        Case 103
            If CapsLock Then
                SendKeys "E"
            Else
                SendKeys "e"
            End If
        Case 104
            If CapsLock Then
                SendKeys "R"
            Else
                SendKeys "r"
            End If
        Case 105
            If CapsLock Then
                SendKeys "T"
            Else
                SendKeys "t"
            End If
        Case 106
            If CapsLock Then
                SendKeys "Y"
            Else
                SendKeys "y"
            End If
        Case 107
            If CapsLock Then
                SendKeys "U"
            Else
                SendKeys "u"
            End If
        Case 108
            If CapsLock Then
                SendKeys "I"
            Else
                SendKeys "i"
            End If
        Case 109
            If CapsLock Then
                SendKeys "O"
            Else
                SendKeys "o"
            End If
        Case 110
            If CapsLock Then
                SendKeys "P"
            Else
                SendKeys "p"
            End If
        Case 111
            If CapsLock Then
                SendKeys "["
            Else
                SendKeys "{{}"
            End If
        Case 112
            If CapsLock Then
                SendKeys "]"
            Else
                SendKeys "{}}"
            End If
        
        Case 201
            If CapsLock Then
                SendKeys "A"
            Else
                SendKeys "a"
            End If
        Case 202
            If CapsLock Then
                SendKeys "S"
            Else
                SendKeys "s"
            End If
        Case 203
            If CapsLock Then
                SendKeys "D"
            Else
                SendKeys "d"
            End If
        Case 204
            If CapsLock Then
                SendKeys "F"
            Else
                SendKeys "f"
            End If
         Case 205
            If CapsLock Then
                SendKeys "G"
            Else
                SendKeys "g"
            End If
        Case 206
            If CapsLock Then
                SendKeys "H"
            Else
                SendKeys "h"
            End If
        Case 207
            If CapsLock Then
                SendKeys "J"
            Else
                SendKeys "j"
            End If
        Case 208
            If CapsLock Then
                SendKeys "K"
            Else
                SendKeys "k"
            End If
        Case 209
            If CapsLock Then
                SendKeys "L"
            Else
                SendKeys "l"
            End If
        Case 210
            If CapsLock Then
                SendKeys ","
            Else
                SendKeys ";"
            End If
        Case 211
            If CapsLock Then
                SendKeys "."
            Else
                SendKeys ":"
            End If
            
        Case 301
            If CapsLock Then
                SendKeys "Z"
            Else
                SendKeys "z"
            End If
        Case 302
            If CapsLock Then
                SendKeys "X"
            Else
                SendKeys "x"
            End If
        Case 303
            If CapsLock Then
                SendKeys "C"
            Else
                SendKeys "c"
            End If
        Case 304
            If CapsLock Then
                SendKeys "V"
            Else
                SendKeys "v"
            End If
         Case 305
            If CapsLock Then
                SendKeys "B"
            Else
                SendKeys "b"
            End If
        Case 306
            If CapsLock Then
                SendKeys "N"
            Else
                SendKeys "n"
            End If
        Case 307
            If CapsLock Then
                SendKeys "M"
            Else
                SendKeys "m"
            End If
        Case 308
            If CapsLock Then
                SendKeys "'"
            Else
                SendKeys """"
            End If
        Case 309
            If CapsLock Then
                SendKeys "?"
            Else
                SendKeys "_"
            End If
        Case 310
            If CapsLock Then
                SendKeys "<"
            Else
                SendKeys "{^}"
            End If
        Case 311
            If CapsLock Then
                SendKeys ">"
            Else
                SendKeys "°"
            End If
        Case 401
            SendKeys " "
    End Select
End Sub

Private Sub KeyArrowSX_Click()
        TextModifica.SetFocus
        SendKeys "{LEFT}"
End Sub

Private Sub KeyArrowDX_Click()
        TextModifica.SetFocus
        SendKeys "{RIGHT}"
End Sub

Private Sub KeyCapsLock_Click()
    TextModifica.SetFocus
    CapsLock = Not CapsLock
    If CapsLock Then
        KeyCapsLock.BackColor = RGB(&HFF, &HFF, &H0) 'giallo
        Maiuscole
    Else
        KeyCapsLock.BackColor = RGB(200, 200, 200)   'grigio
        Minuscole
    End If
End Sub

Private Sub Maiuscole()
        Key(1).Caption = "1"
        Key(2).Caption = "2"
        Key(3).Caption = "3"
        Key(4).Caption = "4"
        Key(5).Caption = "5"
        Key(6).Caption = "6"
        Key(7).Caption = "7"
        Key(8).Caption = "8"
        Key(9).Caption = "9"
        Key(10).Caption = "0"
        Key(11).Caption = "-"
        Key(12).Caption = "/"
        Key(101).Caption = "Q"
        Key(102).Caption = "W"
        Key(103).Caption = "E"
        Key(104).Caption = "R"
        Key(105).Caption = "T"
        Key(106).Caption = "Y"
        Key(107).Caption = "U"
        Key(108).Caption = "I"
        Key(109).Caption = "O"
        Key(110).Caption = "P"
        Key(111).Caption = "["
        Key(112).Caption = "]"
        Key(201).Caption = "A"
        Key(202).Caption = "S"
        Key(203).Caption = "D"
        Key(204).Caption = "F"
        Key(205).Caption = "G"
        Key(206).Caption = "H"
        Key(207).Caption = "J"
        Key(208).Caption = "K"
        Key(209).Caption = "L"
        Key(210).Caption = ","
        Key(211).Caption = "."
        Key(301).Caption = "Z"
        Key(302).Caption = "X"
        Key(303).Caption = "C"
        Key(304).Caption = "V"
        Key(305).Caption = "B"
        Key(306).Caption = "N"
        Key(307).Caption = "M"
        Key(308).Caption = "'"
        Key(309).Caption = "?"
        Key(310).Caption = "<"
        Key(311).Caption = ">"
        

End Sub

Private Sub Minuscole()
        Key(1).Caption = "!"
        Key(2).Caption = "@"
        Key(3).Caption = "£"
        Key(4).Caption = "$"
        Key(5).Caption = "%"
        Key(6).Caption = "&&"
        Key(7).Caption = "*"
        Key(8).Caption = "("
        Key(9).Caption = ")"
        Key(10).Caption = "="
        Key(11).Caption = "+"
        Key(12).Caption = "\"
        Key(101).Caption = "q"
        Key(102).Caption = "w"
        Key(103).Caption = "e"
        Key(104).Caption = "r"
        Key(105).Caption = "t"
        Key(106).Caption = "y"
        Key(107).Caption = "u"
        Key(108).Caption = "i"
        Key(109).Caption = "o"
        Key(110).Caption = "p"
        Key(111).Caption = "{"
        Key(112).Caption = "}"
        Key(201).Caption = "a"
        Key(202).Caption = "s"
        Key(203).Caption = "d"
        Key(204).Caption = "f"
        Key(205).Caption = "g"
        Key(206).Caption = "h"
        Key(207).Caption = "j"
        Key(208).Caption = "k"
        Key(209).Caption = "l"
        Key(210).Caption = ";"
        Key(211).Caption = ":"
        Key(301).Caption = "z"
        Key(302).Caption = "x"
        Key(303).Caption = "c"
        Key(304).Caption = "v"
        Key(305).Caption = "b"
        Key(306).Caption = "n"
        Key(307).Caption = "m"
        Key(308).Caption = """"
        Key(309).Caption = "_"
        Key(310).Caption = "^"
        Key(311).Caption = "°"
        
End Sub



