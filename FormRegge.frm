VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormRegge 
   BackColor       =   &H00C0C0C0&
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
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
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
      Height          =   5655
      Left            =   210
      TabIndex        =   32
      Top             =   1230
      Width           =   5265
      Begin VB.CommandButton Command3 
         Caption         =   "Ticket generator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   1920
         TabIndex        =   46
         Top             =   -1050
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Height          =   2685
         Left            =   60
         TabIndex        =   41
         Top             =   2910
         Width           =   5145
         Begin VB.Label DisplayBundles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   2580
            TabIndex        =   48
            Top             =   750
            Width           =   2205
         End
         Begin VB.Label LabelCartellino 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bundle counter"
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
            Height          =   375
            Index           =   2
            Left            =   2640
            TabIndex        =   47
            Top             =   360
            Width           =   2145
         End
         Begin VB.Label LabelCartellino 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Numero pacco"
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
            Height          =   435
            Index           =   0
            Left            =   180
            TabIndex        =   45
            Top             =   1530
            Width           =   2145
         End
         Begin VB.Label DisplayCartellino 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   150
            TabIndex        =   44
            Top             =   750
            Width           =   2205
         End
         Begin VB.Label DisplayCartellino 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   150
            TabIndex        =   43
            Top             =   1950
            Width           =   2205
         End
         Begin VB.Label LabelCartellino 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Numero barre"
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
            Height          =   435
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   330
            Width           =   2205
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   5
            Left            =   150
            Top             =   1470
            Width           =   2205
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   3
            Left            =   150
            Top             =   270
            Width           =   2205
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   495
            Index           =   4
            Left            =   2580
            Top             =   300
            Width           =   2205
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
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
         Left            =   2550
         TabIndex        =   34
         Top             =   2040
         Width           =   2475
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ticket modify"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2550
         TabIndex        =   33
         Top             =   1080
         Width           =   2475
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0       1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   300
         TabIndex        =   36
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   1635
      End
      Begin VB.Image PrinterSelector 
         Height          =   1155
         Left            =   480
         Top             =   1620
         Width           =   1185
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   240
         Top             =   1260
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   5145
      Left            =   240
      TabIndex        =   23
      Top             =   1230
      Width           =   4545
      Begin VB.Label DisplayCartellino 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1050
         TabIndex        =   31
         Top             =   4440
         Width           =   2205
      End
      Begin VB.Label LabelCartellino 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N. giri fasciatura singola"
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
         Height          =   675
         Index           =   3
         Left            =   1080
         TabIndex        =   30
         Top             =   3840
         Width           =   2115
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   1
         Left            =   2730
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Image Selettore 
         Height          =   1155
         Index           =   0
         Left            =   480
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Fasciatrice"
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
         Height          =   735
         Left            =   990
         TabIndex        =   24
         Top             =   300
         Width           =   2445
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   825
         Index           =   2
         Left            =   960
         Top             =   270
         Width           =   2505
      End
      Begin VB.Label Label01Rif 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0        1"
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
         Left            =   300
         TabIndex        =   27
         Top             =   2100
         Width           =   1635
      End
      Begin VB.Label LblVel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Singola"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   720
         Index           =   5
         Left            =   270
         TabIndex        =   26
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0        1"
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
         Left            =   2520
         TabIndex        =   29
         Top             =   2100
         Width           =   1635
      End
      Begin VB.Label LblVel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Completa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   720
         Index           =   0
         Left            =   2520
         TabIndex        =   28
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   735
         Index           =   7
         Left            =   1050
         Top             =   3720
         Width           =   2205
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   2520
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Shape ShapeRif 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   270
         Top             =   1920
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   6
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   25
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   240
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
         TabIndex        =   13
         Top             =   270
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
         TabIndex        =   12
         Top             =   240
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "FormRegge.frx":0000
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
         TabIndex        =   16
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   -30
      Top             =   1800
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   0
      Top             =   2250
   End
   Begin dp6.ControlloRegge Regge480 
      Height          =   1995
      Left            =   8550
      TabIndex        =   0
      Top             =   6690
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   3519
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -30
      Top             =   4350
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
            Picture         =   "FormRegge.frx":208E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormRegge.frx":2691
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   17
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Straps"
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
      Height          =   3765
      Left            =   11010
      TabIndex        =   18
      Top             =   1230
      Width           =   4095
      Begin VB.CommandButton CommandModificaRegge 
         Caption         =   "MODIFICA"
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
         Left            =   870
         TabIndex        =   19
         Top             =   2820
         Width           =   2595
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Straps"
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
         Left            =   900
         TabIndex        =   22
         Top             =   2190
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Tube length"
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
         Height          =   525
         Left            =   900
         TabIndex        =   21
         Top             =   510
         Width           =   2535
      End
      Begin VB.Label Lbl480Lung 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "99999"
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
         Left            =   870
         TabIndex        =   20
         Top             =   1050
         Width           =   2565
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   705
         Index           =   1
         Left            =   870
         Top             =   2040
         Width           =   2595
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   795
         Index           =   0
         Left            =   870
         Top             =   300
         Width           =   2565
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   40
      Top             =   6540
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8430
      TabIndex        =   39
      Top             =   6540
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   6360
      TabIndex        =   38
      Top             =   1980
      Width           =   1665
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   8070
      TabIndex        =   37
      Top             =   1980
      Width           =   1665
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   2595
      Left            =   7770
      Top             =   6390
      Width           =   165
   End
   Begin VB.Line LineaZero 
      Visible         =   0   'False
      X1              =   8520
      X2              =   8520
      Y1              =   6420
      Y2              =   8940
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
      Left            =   6660
      TabIndex        =   4
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Image Selettore 
      Height          =   1155
      Index           =   2
      Left            =   8280
      Top             =   2760
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
      Left            =   8280
      TabIndex        =   3
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Image Selettore 
      Height          =   1155
      Index           =   3
      Left            =   6660
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Label DisplayPosizione 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "99999"
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
      Left            =   7710
      TabIndex        =   1
      Top             =   5730
      Width           =   1035
   End
   Begin VB.Shape Arcolaio 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   2595
      Left            =   8430
      Top             =   6390
      Width           =   165
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "99999"
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
      Left            =   5880
      TabIndex        =   2
      Top             =   5610
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image Image4 
      Height          =   5400
      Left            =   90
      Picture         =   "FormRegge.frx":2CA3
      Top             =   5010
      Width           =   15030
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   8070
      Top             =   2550
      Width           =   1665
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   6360
      Top             =   2550
      Width           =   1665
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Strapping enable"
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
      Height          =   3015
      Left            =   5910
      TabIndex        =   5
      Top             =   1380
      Width           =   4185
   End
End
Attribute VB_Name = "FormRegge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Lampeggio As Boolean

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
      frmKernel.PaginaCorrente = Index
End Sub

Private Sub Command1_Click()
   frmStoccaggio.Modifica
End Sub

Private Sub Command2_Click()
   With frmStampa
         '.LoadTicketFile "..\Target\Tickets\MairTicket_" & Format(Cartellino.Lingua, "00") & ".TKT"
'         .LoadFixTexts
'         .LengFixTextRefresh (Cartellino.Lingua)
'         .FixVisibleRefresh
'         .RefreshVar
         .PrintExec DB402.Word(18), Cartellino.CampoManuale(2)
    End With
End Sub

Private Sub Command3_Click()
    Shell "..\SouthlandTicket\SouthlandTicket.exe", vbNormalFocus
End Sub

Private Sub DisplayBundles_Click()
    Dim temp As Double
    
    TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par110_Password"))
    TechPasswordForm.Show vbModal
    If TechPasswordForm.LoginSucceeded = False Then Exit Sub
          
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 7000000
    TOUCHNumericPad.ValoreMax = 7999999
    TOUCHNumericPad.Dati = DB402.DWord(40)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        DB402.DWord(40) = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub DisplayCartellino_Click(Index As Integer)
    Dim temp
    
    If Index <> 2 Then Exit Sub
    
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 3
    TOUCHNumericPad.ValoreMax = 6
    TOUCHNumericPad.Dati = DB480.Word(90)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        DB480.Word(90) = TOUCHNumericPad.Dati
        DisplayCartellino(2).caption = TOUCHNumericPad.Dati
    End If
End Sub

Private Sub Lbl480Lung_Click()
  ' legge l'ordine e la ricetta dall'archivio
     
    ModOrder.IDOrdine = DB480.Word(0)
    ModOrder.UploadData DB480.Word(0)
    Ricetta.IDRicetta = ModOrder.IDRicetta
    Ricetta.UploadData ModOrder.IDRicetta
    ' sovrascrive con i dati presenti su plc
    UploadDB480 Ricetta
    ' modifica
    RecipeModifyForm.StrapEnable = True
    RecipeModifyForm.BundleEnable = False
    RecipeModifyForm.SSTab1.TabEnabled(3) = False
    RecipeModifyForm.SSTab1.TabEnabled(4) = False
    RecipeModifyForm.SSTab1.Tab = 1
    RecipeModifyForm.ModZona = True   ' modifica stato zona
    Ricetta.TuboLunghezza = DB480.Word(4) / 1000
    RecipeModifyForm.Show vbModal
    If RecipeModifyForm.OK Then
        ' scrive su plc i dati
        DownloadDB480 Ricetta
        ' salva le modifiche
        'OrdersForm.SaveRecipeData Ricetta
    End If
End Sub

Private Sub PrinterSelector_Click()
    If Param.GetBit("Par112_StampaInserita") Then
        Param.SetBit "Par112_StampaInserita", False
    Else
        Param.SetBit "Par112_StampaInserita", True
    End If
    
    If Param.GetBit("Par112_StampaInserita") Then
        PrinterSelector.Picture = ImageList1.ListImages(2).Picture
        Command1.Visible = True
        Command2.Visible = True
        DB402.Bit(9, 0) = 1
    Else
        PrinterSelector.Picture = ImageList1.ListImages(1).Picture
        Command1.Visible = False
        Command2.Visible = False
        DB402.Bit(9, 0) = 0
    End If
    
End Sub

Private Sub Selettore_Click(Index As Integer)
    Select Case Index
       Case 2
          If DB422.Bit(27, 3) Then
             DB422.Bit(27, 2) = Not DB422.Bit(27, 2)
          Else
             DB422.Bit(27, 2) = True
          End If
       Case 3
          If DB422.Bit(27, 2) Then
             DB422.Bit(27, 3) = Not DB422.Bit(27, 3)
          Else
             DB422.Bit(27, 3) = True
          End If
       End Select
End Sub

Private Sub Timer1_Timer()
       Lampeggio = Not (Lampeggio)
       
       If Not (DB422.Bit(27, 4) And Lampeggio) Then
         '  Shape5(0).BackColor = &HFFFF&
           Arcolaio.BackColor = &HFFFF&
        Else
           Arcolaio.BackColor = &HFF&
         '  Shape5(0).BackColor = &HFF&
        End If
        If Not (DB422.Bit(27, 5) And Lampeggio) Then
          ' Shape5(1).BackColor = &HFFFF&
           Shape6.BackColor = &HFFFF&
        Else
           Shape6.BackColor = &HFF&
         ' Shape5(1).BackColor = &HFF&
        End If
End Sub

Private Sub TimerLocale_Timer()
   Me.Update
End Sub

Public Sub Update()
    Dim i As Integer
    Dim PosizioneInTwips As Long
    Dim DB422W28_change As Boolean
    
    
    ' aggiorna i dati pagina
    DisplayCartellino(2) = DB480.Word(90)
    lblbar(2) = PaginaReggiatura.Ordine_Descrizione
    lblbar(4) = PaginaReggiatura.Ricetta_Descrizione
    
    DisplayBundles = DB402.DWord(40)
    
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    
    If DB422.Bit(27, 2) = False And DB422.Bit(27, 3) = False Then
       DB422.Bit(27, 2) = True
    End If
    
    Lbl480Lung.caption = Conv_UM.Conversione(DB480.Word(4), UM.mm, UM.inch, 1)
    If DB422.DatiCambiati Then
       DB422W28_change = True
       DB422.DatiCambiati = False
    Else
       DB422W28_change = False
       DB422.DatiCambiati = False
    End If
      
      If DB422.Bit(27, 2) Then
           Selettore(2).Picture = ImageList1.ListImages(2).Picture
        Else
            Selettore(2).Picture = ImageList1.ListImages(1).Picture
        End If
        If DB422.Bit(27, 3) Then
            Selettore(3).Picture = ImageList1.ListImages(2).Picture
        Else
            Selettore(3).Picture = ImageList1.ListImages(1).Picture
        End If

    If DB480.DatiCambiati Or DB422W28_change Then
        
        Regge480.VisualizzaLabelLunghezza = True
        Regge480.VisualizzaQuote = True
        Regge480.PaccoLunghezza = DB480.Word(4) / 1000
        For i = 1 To 12
            Regge480.QuotaReggia(i) = DB480.Word(64 + i * 2) / 1000
        Next i
        Regge480.Refresh
        
        ' fasciatura ==========
        If DB480.Bit(62, 1) Then
            Selettore(0).Picture = ImageList1.ListImages(2).Picture
        Else
            Selettore(0).Picture = ImageList1.ListImages(1).Picture
        End If
        
        If DB480.Bit(62, 2) Then
            Selettore(1).Picture = ImageList1.ListImages(2).Picture
        Else
            Selettore(1).Picture = ImageList1.ListImages(1).Picture
        End If
        '=======================
        ' calcolo posizione in reggiatura
        PosizioneInTwips = 0
        If DB480.Word(4) > 0 Then
            PosizioneInTwips = LineaZero.X1 - (DB422.Word(28) * (Regge480.Width - 100) / DB480.Word(4))
        End If
        If PosizioneInTwips > LineaZero.X1 Then PosizioneInTwips = LineaZero.X1
        If PosizioneInTwips < 0 Then PosizioneInTwips = 0
        Regge480.Left = PosizioneInTwips
        DB480.DatiCambiati = False
    End If
    DisplayCartellino(0).caption = DB402.Word(12)   ' numero pacco
    DisplayCartellino(1).caption = DB402.Word(14)   ' numero tubi
    DisplayPosizione.caption = Conv_UM.Conversione(DB422.Word(28) / 1000#, UM.mt, UM.inch)
End Sub

Private Sub Form_Load()

 '  Picture = LoadPicture("..\bitmap\SfondoDP6_ver2_0.jpg")
   WindowState = 2
   
   ScritteMultilingua
End Sub


Private Sub Form_Activate()
    Dim OrdineCorrente As OrderClass
    
    Me.Update
    TimerLocale.Enabled = True
    TimerLocale.Interval = 250
    Barra21.Selezionato = 5
    
    ' aggiornamento dati ordine attuale
    Set OrdineCorrente = New OrderClass
    OrdineCorrente.IDOrdine = DB480.Word(0)
    If OrdersForm.LoadOrderData(OrdineCorrente) Then
        'DisplayCodice.Caption = OrdineCorrente.Descrizione
    End If
      ' abilitazione temporizzatore locale
    'Timer1.Enabled = True
    ' (disattivare in Form_deactivate)
    
    ' abilita la fasciatura
    
    If Param.GetNumber("Par213_Fasciatura") = 1 Then
       Shape2(2).Visible = True
    '   OptionFasciatura(0).Visible = True
    '   OptionFasciatura(1).Visible = True
    '   OptionFasciatura(2).Visible = True
    '   OptionFasciatura(3).Visible = True
       Label5.Visible = True
       Frame3.Visible = True
    Else
       Shape2(2).Visible = False
     '  OptionFasciatura(0).Visible = False
     '  OptionFasciatura(1).Visible = False
     '  OptionFasciatura(2).Visible = False
     '  OptionFasciatura(3).Visible = False
       Label5.Visible = False
       Frame3.Visible = False
    End If
        RecipeModifyForm.SSTab1.Tab = 1
        
    If Param.GetBit("Par112_StampaInserita") Then
        PrinterSelector.Picture = ImageList1.ListImages(2).Picture
        Command1.Visible = True
        Command2.Visible = True
        DB402.Bit(9, 0) = 1
    Else
        PrinterSelector.Picture = ImageList1.ListImages(1).Picture
        Command1.Visible = False
        Command2.Visible = False
        DB402.Bit(9, 0) = 0
    End If
        ScritteMultilingua
        Me.Update
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
End Sub

'Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
'   frmKernel.PaginaCorrente = Index
'End Sub

Private Sub CommandModificaRegge_Click()
    ' legge l'ordine e la ricetta dall'archivio
    ModOrder.IDOrdine = DB480.Word(0)
    OrdersForm.LoadOrderData ModOrder
    Ricetta.IDRicetta = ModOrder.IDRicetta
    OrdersForm.LoadRecipeData Ricetta
    ' sovrascrive con i dati presenti su plc
    UploadDB480 Ricetta
    ' modifica
    RecipeModifyForm.StrapEnable = True
    RecipeModifyForm.BundleEnable = False
    RecipeModifyForm.SSTab1.TabEnabled(3) = False
    RecipeModifyForm.SSTab1.TabEnabled(4) = False
    RecipeModifyForm.SSTab1.Tab = 1
    RecipeModifyForm.Show vbModal
    If RecipeModifyForm.OK Then
        ' scrive su plc i dati
        DownloadDB480 Ricetta
        ' salva le modifiche
        OrdersForm.SaveRecipeData Ricetta
    End If
End Sub

Sub ScritteMultilingua()
    LabelCartellino(0).caption = Param.Text("Numero pacco")
    LabelCartellino(1).caption = Param.Text("Ntubi")
    Label5.caption = Param.Text("Fasciatrice")
    Label3.caption = Param.Text("Straps")
    CommandModificaRegge.caption = Param.Text("MODIFICA")
    lblbar(5) = Param.Text("Straps page")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
    Label8.caption = Param.Text("LungTubo")
    lblbar(1) = Param.Text("Pagina")
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
                .NomeFile = "Reggiatura_pagina.htm"
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
