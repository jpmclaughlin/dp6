VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PrinterForm 
   BackColor       =   &H80000005&
   Caption         =   "."
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   12450
   WindowState     =   1  'Minimized
   Begin VB.Frame StampaFrame 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9795
      Left            =   5580
      TabIndex        =   44
      Top             =   60
      Width           =   4755
      Begin VB.PictureBox PrinterSelector 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   1635
         Picture         =   "PrinterForm.frx":0000
         ScaleHeight     =   1200
         ScaleWidth      =   1350
         TabIndex        =   47
         Top             =   1965
         Width           =   1350
      End
      Begin VB.CommandButton F9Command 
         Caption         =   "F9 - Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   900
         MaskColor       =   &H0000FFFF&
         TabIndex        =   46
         Top             =   6060
         UseMaskColor    =   -1  'True
         Width           =   3000
      End
      Begin VB.CommandButton F12Command 
         Caption         =   "F12 - Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1300
         Left            =   900
         TabIndex        =   45
         Top             =   7530
         Width           =   3000
      End
      Begin MSCommLib.MSComm PrinterMSComm 
         Left            =   300
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InBufferSize    =   512
         OutBufferSize   =   2048
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
         Left            =   1500
         TabIndex        =   49
         Top             =   1485
         Width           =   1635
      End
      Begin VB.Label PrinterDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   48
         Top             =   840
         Width           =   3600
      End
   End
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
      Height          =   6855
      Left            =   60
      TabIndex        =   9
      Top             =   50
      Width           =   5475
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   42
         Text            =   "8"
         Top             =   2304
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   0
         Left            =   180
         MaxLength       =   20
         TabIndex        =   41
         Text            =   "0"
         Top             =   60
         Width           =   5160
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   40
         Text            =   "1"
         Top             =   780
         Width           =   5160
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   39
         Text            =   "2"
         Top             =   1140
         Width           =   5160
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   38
         Text            =   "32"
         Top             =   6420
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   37
         Text            =   "31"
         Top             =   6420
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   36
         Text            =   "30"
         Top             =   6066
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   35
         Text            =   "29"
         Top             =   6060
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   34
         Text            =   "28"
         Top             =   5724
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   33
         Text            =   "27"
         Top             =   5730
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   32
         Text            =   "26"
         Top             =   5382
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   31
         Text            =   "25"
         Top             =   5385
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         BorderStyle     =   0  'None
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
         TabIndex        =   29
         Text            =   "23"
         Top             =   5040
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   28
         Text            =   "22"
         Top             =   4698
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   27
         Text            =   "21"
         Top             =   4695
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   26
         Text            =   "20"
         Top             =   4356
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   25
         Text            =   "19"
         Top             =   4350
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   24
         Text            =   "18"
         Top             =   4014
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   23
         Text            =   "17"
         Top             =   4020
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   22
         Text            =   "16"
         Top             =   3672
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   21
         Text            =   "15"
         Top             =   3675
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   20
         Text            =   "14"
         Top             =   3330
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   19
         Text            =   "13"
         Top             =   3330
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   18
         Text            =   "12"
         Top             =   2988
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   17
         Text            =   "11"
         Top             =   2985
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   16
         Text            =   "10"
         Top             =   2646
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   15
         Text            =   "9"
         Top             =   2640
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   14
         Text            =   "7"
         Top             =   2310
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   13
         Text            =   "6"
         Top             =   1962
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   12
         Text            =   "4"
         Top             =   1620
         Width           =   3375
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   11
         Text            =   "5"
         Top             =   1965
         Width           =   1755
      End
      Begin VB.TextBox PrintEdit 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         TabIndex        =   10
         Text            =   "3"
         Top             =   1620
         Width           =   1755
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   5400
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Shape Shape1 
         Height          =   6795
         Left            =   60
         Top             =   0
         Width           =   5355
      End
   End
   Begin VB.Frame WeightFrame 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9825
      Left            =   10425
      TabIndex        =   0
      Top             =   60
      Width           =   4600
      Begin VB.PictureBox WeightSelector 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1200
         Left            =   1470
         Picture         =   "PrinterForm.frx":5542
         ScaleHeight     =   1200
         ScaleWidth      =   1350
         TabIndex        =   8
         Top             =   1860
         Width           =   1350
      End
      Begin VB.CommandButton WeightResetCommand 
         BackColor       =   &H0000FF00&
         Caption         =   ">0<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3915
         Width           =   1935
      End
      Begin MSCommLib.MSComm WeightMSComm 
         Left            =   1320
         Top             =   3960
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         BaudRate        =   4800
         ParitySetting   =   2
         DataBits        =   7
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
         Left            =   1335
         TabIndex        =   43
         Top             =   1455
         Width           =   1635
      End
      Begin VB.Label LastBundleWeightDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   480
         TabIndex        =   7
         Top             =   6810
         Width           =   3600
      End
      Begin VB.Label WeightSumDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   495
         TabIndex        =   3
         Top             =   8490
         Width           =   3600
      End
      Begin VB.Label WeightDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   3600
      End
      Begin VB.Label CurrentWeightLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Current weight"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label LastBundleWeightLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Last bundle weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   360
         TabIndex        =   5
         Top             =   6345
         Width           =   3855
      End
      Begin VB.Label WeightSumLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current order weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   375
         TabIndex        =   4
         Top             =   7905
         Width           =   3840
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13080
      Top             =   10080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   90
      ImageHeight     =   80
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrinterForm.frx":AA84
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrinterForm.frx":FFD6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PrinterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'********************************************************
'       Variabili per stampante
'********************************************************

Public PrinterCommEnabled As Boolean

' variabili comuni a pesa e stampante per comunicazione seriale
Private SOH As String
Private STX As String
Private CR As String
Private NUL As String
Private ETX As String
Private EOT As String
Private ENQ As String
Private ACK As String
Private POL As String
Private OnLineString As String
Private NotEnabledString As String
Private Const Timeout As Integer = 1

'********************************************************
'       Variabili per pesa
'********************************************************
Private Enum TWeightState
    WInit = 10
    WStandby = 11
    WGrossRead = 31
    WErrorRead = 33
    WResetRead = 35
    WError = 40
    WPause = 41
End Enum
Private WeightState As TWeightState
Private WeightDelayCounter As Integer
Private ActualWeight As Double
Private ValidWeightFlag As Boolean
Private WeightResetFlag As Boolean
Dim WeightRequestString() As Byte
Dim WeightResetString() As Byte
Dim WeightErrorString() As Byte
Dim WeightReceivedData() As Byte
Private Const InBufferSize As Integer = 64



Private Sub Form_Load()
'    ' testi frame contatori
'    ProductionFrame.Caption = Param.Text("Production")
'    CurrentBundleLabel.Caption = Param.Text("BundleNum")
'    CurrentTubesLabel.Caption = Param.Text("BundleTubes")

'    ' testi dimensioni tubo
'    TubeFrame.Caption = Param.Text("TubeData") & Unit.mmString
'    DimensionLabel.Caption = Param.Text("TubeDimension")
'    DiameterLabel.Caption = Param.Text("Diameter")
'    ThickLabel.Caption = Param.Text("Thickness")
'    LengthLabel.Caption = Param.Text("TubeLength")

    ' altri testi
'    StampaFrame.Caption = Param.Text("Txt007") ' stampa
'    WeightFrame.Caption = Param.Text("Txt086")  ' pesa
'    WeightSumLabel.Caption = Param.Text("OrderWeight") & Unit.KgString
'    CurrentWeightLabel.Caption = Param.Text("CurrentWeight") & Unit.KgString
'    LastBundleWeightLabel.Caption = Param.Text("LastBundleWeight") & Unit.KgString

'    F9Command.Caption = Param.Text("Txt001")    ' modifica
'    F12Command.Caption = Param.Text("Txt007")   ' stampa
    ' *************** stampante ***********
'    StampaFrame.Visible = Param.GetBit("Par114_PaginaStampa")
'    PrinterCommEnabled = Param.GetBit("Par112_StampaInserita")
    If PrinterCommEnabled Then
        PrinterSelector.Picture = ImageList1.ListImages(2).Picture
    Else
        PrinterSelector.Picture = ImageList1.ListImages(1).Picture
    End If

    ' prepara stringhe predefinite
    SOH = Chr$(&H1)
    STX = Chr$(&H2)
    CR = Chr$(&HD)
    NUL = Chr(&H0)
    ETX = Chr(&H3)
    EOT = Chr(&H4)
    ENQ = Chr(&H5)
    ACK = Chr(&H6)
    POL = Chr(&H70)

'    OnLineString = Param.Text("Txt025") ' on line
'    NotEnabledString = Param.Text("Txt023") ' non abilitata
    


    ' inizializza variabili
    ' ************ pesa ****************
'    WeightFrame.Visible = Param.GetBit("Par113_PaginaPesatura")
    ' se non c'è la stampante il frame della pesa viene allineato a destra
'    If Param.GetBit("Par114_PaginaStampa") = False Then
'        WeightFrame.Width = WeightFrame.Left + WeightFrame.Width
'        WeightFrame.Left = LabelFrame.Width + LabelFrame.Left
'    End If
    

    WeightState = WInit

    ' preparazione stringa richiesta peso
    ReDim WeightRequestString(7)
    WeightRequestString(0) = &H1   ' Slave address
    WeightRequestString(1) = &H3   ' function code : Read register
    WeightRequestString(2) = &H0   ' register address - HightByte
    WeightRequestString(3) = &HB   ' register address - LowByte
    WeightRequestString(4) = &H0   ' register length - HightByte
    WeightRequestString(5) = &H3   ' register length - LowByte
    WeightRequestString(6) = &H74  ' CRC LowByte
    WeightRequestString(7) = &H9   ' CRC HightByte

    ' preparazione stringa azzeramento pesa
    ReDim WeightResetString(7)
    WeightResetString(0) = &H1   ' Slave address
    WeightResetString(1) = &H6   ' function code : write register
    WeightResetString(2) = &H0   ' register address - HightByte
    WeightResetString(3) = &H1D  ' register address - LowByte
    WeightResetString(4) = &H0   ' register length - HightByte
    WeightResetString(5) = &H8   ' register length - LowByte
    WeightResetString(6) = &H18  ' CRC LowByte
    WeightResetString(7) = &HA   ' CRC HightByte

    ' preparazione stringa lettura codice di errore
    ReDim WeightErrorString(7)
    WeightErrorString(0) = &H1   ' Slave address
    WeightErrorString(1) = &H3   ' function code : Read register
    WeightErrorString(2) = &H0   ' register address - HightByte
    WeightErrorString(3) = &H8   ' register address - LowByte
    WeightErrorString(4) = &H0   ' register length - HightByte
    WeightErrorString(5) = &H1   ' register length - LowByte
    WeightErrorString(6) = &H5   ' CRC LowByte
    WeightErrorString(7) = &HC8  ' CRC HightByte



End Sub

'**************************************************
' Funzione per impedire la chiusura del form da parte
' dell'utente
'**************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormControlMenu Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub



Public Sub Update()
        
    If PrinterCommEnabled Then
        PrinterDisplay.Caption = OnLineString
    Else
        PrinterDisplay.Caption = NotEnabledString
    End If


    '*************************************************************
    '  Dfa pesa NOBEL
    '*************************************************************
    Select Case WeightState
    '0) Init
    Case WInit
        WeightState = WStandby
        If DB402.Bit(8, 4) Then
            ' ridimensiona buffer byte alla stessa dimensione del buffer seriale
            WeightMSComm.InBufferSize = InBufferSize
            ReDim WeightReceivedData(InBufferSize)
            ' chiude porta
            If WeightMSComm.PortOpen Then WeightMSComm.PortOpen = False
            ' setta parametri porta
            WeightMSComm.InputMode = comInputModeBinary
            WeightMSComm.Settings = "9600,N,8,1"
            On Error Resume Next
                WeightMSComm.CommPort = Param.GetNumber("Par115_SerialePesa")
                If Err.Number <> 0 Then MsgBox "Parameter ""Par115_SerialePesa"" error", vbOKOnly
            On Error GoTo 0
            On Error Resume Next
                ' riapre porta
                If Not WeightMSComm.PortOpen Then WeightMSComm.PortOpen = True
                ' se non riesce ad aprirla significa che la porta non è disponibile
                If Err.Number <> 0 Then
                    MainMDIForm.PcAlarm(WeightCommPort) = True
                    WeightDisplay.Caption = Param.Text("Txt022") ' No porta seriale
                Else
                    MainMDIForm.PcAlarm(WeightCommPort) = False
                    WeightMSComm.Output = WeightErrorString
                    WeightState = WErrorRead
                End If
            On Error GoTo 0
        Else
            On Error Resume Next
            If WeightMSComm.PortOpen Then WeightMSComm.PortOpen = False
            WeightDisplay.Caption = Param.Text("Txt023") ' Non abilitata
            MainMDIForm.PcAlarm(WeightCommPort) = False
            MainMDIForm.PcAlarm(WeightCommErr) = False
            MainMDIForm.PcAlarm(WeightFault) = False
        End If
        ValidWeightFlag = False
        WeightResetFlag = False

    Case WStandby   ' do nothing state if Weight is not enabled

    Case WErrorRead
        If WeightMSComm.InBufferCount >= 7 Then
            WeightReceivedData = WeightMSComm.Input
            If WeightReceivedData(4) <> 0 Then
                WeightDisplay.Caption = "Error " & WeightReceivedData(4)
                ValidWeightFlag = False
                WeightResetFlag = False
                MainMDIForm.PcAlarm(WeightFault) = True
                WeightState = WPause
            Else
                WeightMSComm.Output = WeightRequestString
                WeightState = WGrossRead
            End If
            WeightDelayCounter = 0
        Else
            WeightDelayCounter = WeightDelayCounter + 1
            If WeightDelayCounter > Timeout Then WeightState = WError
        End If


    Case WGrossRead  ' lettura peso
        If WeightMSComm.InBufferCount >= 11 Then
            WeightReceivedData = WeightMSComm.Input
            ActualWeight = WeightDecode
            WeightDisplay.Caption = Unit.kg_To_Display_kg(ActualWeight)
            MainMDIForm.PcAlarm(WeightFault) = False
            MainMDIForm.PcAlarm(WeightCommErr) = False
            WeightDelayCounter = 0
            If DB402.Bit(8, 8) Then
                DB402.Bit(8, 8) = False
                WeightMSComm.Output = WeightResetString
                WeightState = WResetRead
            Else
                WeightMSComm.Output = WeightErrorString
                WeightState = WErrorRead
            End If
        Else
            WeightDelayCounter = WeightDelayCounter + 1
            If WeightDelayCounter > Timeout Then WeightState = WError
        End If


    Case WResetRead
        If WeightMSComm.InBufferCount >= 8 Then
            WeightReceivedData = WeightMSComm.Input
            WeightDelayCounter = 0
            WeightMSComm.Output = WeightErrorString
            WeightState = WErrorRead
        Else
            WeightDelayCounter = WeightDelayCounter + 1
            If WeightDelayCounter > Timeout Then WeightState = WError
        End If


    '9) Errore
    Case WError
        WeightDelayCounter = 0
        WeightResetFlag = False
        ValidWeightFlag = False
        MainMDIForm.PcAlarm(WeightCommErr) = True
        MainMDIForm.PcAlarm(WeightFault) = False
        WeightDisplay.Caption = Param.Text("Txt024") ' non collegata
        WeightState = WPause

    '10) Pausa
    Case WPause
        WeightDelayCounter = WeightDelayCounter + 1
        If WeightDelayCounter > 0 Then
            WeightDelayCounter = 0
            WeightMSComm.Output = WeightErrorString
            WeightState = WErrorRead
        End If
    Case Else
        WeightState = WPause
    End Select

    ' selettore di pesatura
    If DB402.Bit(8, 4) Then
        WeightSelector.Picture = ImageList1.ListImages(2).Picture
    Else
        WeightSelector.Picture = ImageList1.ListImages(1).Picture
    End If

End Sub

' decodifica il peso nell'array di bytes ricevuti
Private Function WeightDecode() As Double
    Dim tmp As Byte
    If (WeightReceivedData(3) And &H80) <> 0 Then
        ' conversione peso negativo
        WeightDecode = 0
        tmp = Not (WeightReceivedData(3))
        WeightDecode = WeightDecode + tmp * 16777216#
        tmp = Not (WeightReceivedData(4))
        WeightDecode = WeightDecode + tmp * 65536#
        tmp = Not (WeightReceivedData(5))
        WeightDecode = WeightDecode + tmp * 256#
        tmp = Not (WeightReceivedData(6))
        WeightDecode = WeightDecode + tmp
        WeightDecode = -WeightDecode - 1#
    Else
        ' Conversione peso positivo
        WeightDecode = WeightDecode + WeightReceivedData(3) * 16777216#
        WeightDecode = WeightDecode + WeightReceivedData(4) * 65536#
        WeightDecode = WeightDecode + WeightReceivedData(5) * 256#
        WeightDecode = WeightDecode + WeightReceivedData(6)
    End If

    ' Numero di decimali
    Select Case WeightReceivedData(8)
        Case 1
            WeightDecode = WeightDecode / 10#
        Case 2
            WeightDecode = WeightDecode / 100#
        Case 3
            WeightDecode = WeightDecode / 1000#
        Case 4
           WeightDecode = WeightDecode / 10000#
        Case 5
            WeightDecode = WeightDecode / 100000#
    End Select
End Function

' pulsante azzeramento pesa
Private Sub WeightResetCommand_Click()
    DB402.Bit(8, 8) = True
End Sub

' pulsante abilitazione pesa
Private Sub WeightSelector_Click()
    ' aggiorna flag di pesa abilitata
    DB402.Bit(8, 4) = Not DB402.Bit(8, 4)
    ' fa ripartire il dfa di comunicazione
    WeightState = WInit
End Sub

Private Sub PrinterSelector_Click()
    If PrinterSelector.Picture = ImageList1.ListImages(1).Picture Then
        PrinterSelector.Picture = ImageList1.ListImages(2).Picture
        PrinterCommEnabled = True
    Else
        PrinterSelector.Picture = ImageList1.ListImages(1).Picture
        PrinterCommEnabled = False
    End If
    ' ricopia valore su parametro nel file dei parametri
    Param.SetBit "Par112_StampaInserita", PrinterCommEnabled
    ' fa ripartire il dfa di comunicazione

End Sub



  


'********************************************************
' Funzioni di interfaccia pubblica per pesa e stampante
'********************************************************


' Aggiorna campi fissi del cartellino (in quelli variabili vengono cariacati i dati del pacco precedente)
Public Sub FixedLabelUpdate()
'    Dim TotString As String
'    Dim i As Integer
'    On Error Resume Next
'    ' carica i dati dell'ordine attualmente in zona pesa
'    If OrdersForm.LoadOrderData(WeightOrder) Then
'        TotString = WeightOrder.PrinterData
'        ' trasferimento al video dei dati aggiornati
'        For i = 0 To (PrintFieldNumber - 1)
'            PrintEdit(i).Text = RTrim(Mid(TotString, 1 + PrintFieldLength * i, PrintFieldLength))
'        Next i
'    End If
End Sub

' Funzione per stampa etichetta vecchia (prelevata da archivio)
Public Sub OldLabelPrint(OldData As String)
    Dim i As Integer
    On Error Resume Next
    ' trasferimento al video dei dati aggiornati
    For i = 0 To (PrintFieldNumber - 1)
        PrintEdit(i).Text = RTrim(Mid(OldData, 1 + PrintFieldLength * i, PrintFieldLength))
    Next i
    PrintExec
End Sub


' Aggiorna campi fissi evariabili variabili (peso, numero pacco e tubi)
Public Sub VariableLabelUpdate()

    PrintEdit(4).Text = Cartellino.CampoManuale(1)
    PrintEdit(6).Text = Cartellino.CampoManuale(2)
    PrintEdit(8).Text = Cartellino.CampoManuale(3)
    PrintEdit(10).Text = Cartellino.CampoManuale(4)
    PrintEdit(12).Text = Cartellino.CampoManuale(5)
    PrintEdit(14).Text = Cartellino.CampoManuale(6)
    PrintEdit(16).Text = Cartellino.CampoManuale(7)
    PrintEdit(18).Text = Cartellino.CampoManuale(8)
    PrintEdit(20).Text = Cartellino.CampoManuale(9)
    PrintEdit(22).Text = Cartellino.Data
    PrintEdit(24).Text = Cartellino.NumeroPacco
    PrintEdit(26).Text = Cartellino.NumeroTubi
    PrintEdit(28).Text = Cartellino.PesoPacco
    PrintEdit(30).Text = ""
    PrintEdit(32).Text = ""
        
End Sub




' comando di stampa
Public Sub PrintRequest()
    PrintExec
End Sub

' stato pesa
Public Function ValidWeight() As Boolean
    If ValidWeightFlag Then
        ValidWeight = True
    Else
        ValidWeight = False
    End If
    If Not DB402.Bit(8, 4) Then ValidWeight = True
End Function


' modifica dati
Public Sub F9Command_Click()
'    Dim TmpOrder As OrderClass
'    Set TmpOrder = New OrderClass
'
'    TmpOrder.PlcCode = WeightOrder.PlcCode
'    OrdersForm.LoadOrderData TmpOrder
'    ' trasferisce il puntatore del pacco alla finestra di modifica dati
'    ModifyForm.GetData TmpOrder, True, 0
'    ' esegue la finestra modifica dati
'    ModifyForm.Show (vbModal)
'    If ModifyForm.Ok Then
'        ' salva su database la ricetta modificata
'        OrdersForm.SaveOrderData TmpOrder, False
'        ' trasmette la nuova ricetta a tutte le mappe plc
'        OrdersForm.SendModifiedData
'    End If
'    ' aggiorna l'immagine
'    FixedLabelUpdate
End Sub

' print command
Public Sub F12Command_Click()
    PrintExec
End Sub


    Private Sub PrintExec()
        Dim i As Integer
        Dim HideFlag As Boolean
        If PrinterCommEnabled Then
            '0) Si ridimensiona il form per far visualizzare tutta l'etichetta anche
            '   se il form non è ancora massimizzato (evita stampa troncata)
            On Error Resume Next
                ' Rende visiblile il frame se è nascosto
                If PrinterForm.WindowState <> vbMaximized Then
                    PrinterForm.Visible = True
                    PrinterForm.ZOrder (0)
                    PrinterForm.WindowState = vbMaximized
                    HideFlag = True
                Else
                    HideFlag = False
                End If

                If Me.Height < (LabelFrame.Height + LabelFrame.Top) Then
                    Me.Height = LabelFrame.Height + LabelFrame.Top
                End If
                '1) rende invisibile la parte del form che non bisogna stampare
                WeightFrame.Visible = False
                StampaFrame.Visible = False
                
                ' 2) trasferisce ai PrintData le stringhe prelevate dai PrintEdit
                '    perchè ci sono problemi di font a stampare direttamente i
                '    PrintEdit
'                For i = 0 To PrintFieldNumber - 1
'                    PrintData(i).Caption = PrintEdit(i).Text
'                    PrintData(i).Visible = True
'                    PrintEdit(i).Visible = False
'                Next i

                ' 3) Comando di stampa
                Me.PrintForm
                
                ' 4) Ripristina la visibilità dei PrintEdit e del resto del Form
'                For i = 0 To PrintFieldNumber - 1
'                    PrintData(i).Visible = False
'                    PrintEdit(i).Visible = True
'               Next i
                WeightFrame.Visible = Param.GetBit("Par113_PaginaPesatura")
                StampaFrame.Visible = Param.GetBit("Par114_PaginaStampa")
                ' nasconde il form se era nascosto in precedenza
                If HideFlag Then PrinterForm.ZOrder (1)
            On Error GoTo 0
       End If
    End Sub

