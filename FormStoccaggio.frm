VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStoccaggio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   9330
      TabIndex        =   50
      Top             =   1170
      Width           =   2805
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   11880
      TabIndex        =   49
      Top             =   1170
      Width           =   3255
   End
   Begin VB.Frame FramePrinter 
      BackColor       =   &H00C0C0C0&
      Height          =   4875
      Left            =   11880
      TabIndex        =   42
      Top             =   1260
      Width           =   3135
      Begin VB.CommandButton CommandPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   660
         TabIndex        =   44
         Top             =   4050
         Width           =   1980
      End
      Begin VB.CommandButton CommandModifica 
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   660
         MaskColor       =   &H0000FFFF&
         TabIndex        =   43
         Top             =   3180
         UseMaskColor    =   -1  'True
         Width           =   1980
      End
      Begin VB.Frame FrameStampa 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   60
         TabIndex        =   47
         Top             =   150
         Width           =   2985
      End
      Begin VB.Image PrinterSelector 
         Height          =   1155
         Left            =   1050
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Printer   "
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
         Left            =   780
         TabIndex        =   46
         Top             =   630
         Width           =   1635
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
         Index           =   0
         Left            =   840
         TabIndex        =   45
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   780
         Top             =   1050
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   4875
      Left            =   2040
      TabIndex        =   34
      Top             =   1260
      Width           =   4755
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Time maximum drying"
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
         Height          =   465
         Left            =   240
         TabIndex        =   53
         Top             =   4110
         Width           =   2535
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   52
         Top             =   4020
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3090
         TabIndex        =   51
         Top             =   4050
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   41
         Top             =   3090
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scoli"
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
         Height          =   525
         Left            =   1530
         TabIndex        =   40
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label LabelPesa 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pesa"
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
         TabIndex        =   39
         Top             =   570
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
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
         Left            =   360
         TabIndex        =   38
         Top             =   1050
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image WeightSelector 
         Height          =   1155
         Left            =   240
         Top             =   1410
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
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
         Index           =   2
         Left            =   1800
         TabIndex        =   37
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Image SelScoli 
         Height          =   1155
         Left            =   1740
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3090
         TabIndex        =   36
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Time minimum drying"
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
         Left            =   210
         TabIndex        =   35
         Top             =   3180
         Width           =   2625
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   5
         Left            =   150
         Top             =   3120
         Width           =   2955
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   585
         Index           =   4
         Left            =   1530
         Top             =   480
         Width           =   1635
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   585
         Index           =   6
         Left            =   90
         Top             =   480
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   1530
         Top             =   1050
         Width           =   1635
      End
      Begin VB.Shape ShapeRif 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1635
         Left            =   90
         Top             =   990
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   150
         Top             =   4050
         Width           =   2955
      End
   End
   Begin VB.Frame FrameSoloScoli 
      BackColor       =   &H00C0C0C0&
      Height          =   4875
      Left            =   2040
      TabIndex        =   26
      Top             =   1230
      Width           =   4755
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Portatavolette"
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
         Height          =   705
         Left            =   2760
         TabIndex        =   33
         Top             =   570
         Width           =   1590
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
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
         Index           =   4
         Left            =   3030
         TabIndex        =   32
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Image Image5 
         Height          =   1155
         Left            =   2940
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   31
         Top             =   3330
         Width           =   315
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
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
         Left            =   3180
         TabIndex        =   30
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ritardo pesa"
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
         Left            =   300
         TabIndex        =   29
         Top             =   3450
         Width           =   4065
      End
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   600
         Top             =   1380
         Width           =   1185
      End
      Begin VB.Label Label01 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0      1"
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
         Index           =   3
         Left            =   660
         TabIndex        =   28
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scoli"
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
         Left            =   300
         TabIndex        =   27
         Top             =   570
         Width           =   1770
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   10
         Left            =   270
         Top             =   3330
         Width           =   2925
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   525
         Index           =   9
         Left            =   2580
         Top             =   570
         Width           =   1875
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   525
         Index           =   8
         Left            =   270
         Top             =   570
         Width           =   1815
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   2580
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   270
         Top             =   1020
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   4875
      Left            =   9330
      TabIndex        =   18
      Top             =   1140
      Width           =   2685
      Begin VB.CommandButton WeightResetCommand 
         Caption         =   ">0<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1590
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label WeightDisplay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   180
         TabIndex        =   25
         Top             =   840
         Width           =   2190
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
         Index           =   3
         Left            =   180
         TabIndex        =   24
         Top             =   3750
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
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   3150
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
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   4200
         Width           =   2205
      End
      Begin VB.Label LabelCartellino 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PESO Reale"
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
         Height          =   630
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   210
         Width           =   2205
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
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   2730
         Width           =   2145
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   7
         Left            =   180
         Top             =   3720
         Width           =   2205
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   675
         Index           =   2
         Left            =   180
         Top             =   180
         Width           =   2205
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   3
         Left            =   180
         Top             =   2670
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   48
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   2415
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image4 
         Height          =   1050
         Left            =   3180
         Picture         =   "FormStoccaggio.frx":0000
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
         TabIndex        =   10
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   270
      Top             =   2100
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   270
      Top             =   2640
   End
   Begin MSCommLib.MSComm WeightMSComm 
      Left            =   150
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   4800
      ParitySetting   =   2
      DataBits        =   7
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   690
      Top             =   1230
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
            Picture         =   "FormStoccaggio.frx":208E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStoccaggio.frx":2691
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Label DisplayPacchi 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
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
      Left            =   7770
      TabIndex        =   17
      Top             =   7080
      Width           =   675
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   12150
      X2              =   12150
      Y1              =   8640
      Y2              =   7620
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   12930
      X2              =   12930
      Y1              =   8640
      Y2              =   7620
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   11910
      X2              =   13050
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   12030
      X2              =   12150
      Y1              =   7740
      Y2              =   7680
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   12960
      X2              =   13080
      Y1              =   7680
      Y2              =   7740
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   12150
      X2              =   12030
      Y1              =   7680
      Y2              =   7620
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   12960
      X2              =   13080
      Y1              =   7680
      Y2              =   7620
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10620
      X2              =   10620
      Y1              =   8640
      Y2              =   7620
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10740
      X2              =   10740
      Y1              =   8640
      Y2              =   7620
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10380
      X2              =   11040
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10500
      X2              =   10620
      Y1              =   7740
      Y2              =   7680
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10740
      X2              =   10860
      Y1              =   7680
      Y2              =   7740
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10620
      X2              =   10500
      Y1              =   7680
      Y2              =   7620
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      Index           =   1
      X1              =   10740
      X2              =   10860
      Y1              =   7680
      Y2              =   7620
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   6810
      X2              =   7050
      Y1              =   8220
      Y2              =   7920
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   7050
      X2              =   7830
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   8310
      X2              =   9090
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   9090
      X2              =   9300
      Y1              =   7920
      Y2              =   8250
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   7830
      X2              =   8070
      Y1              =   7920
      Y2              =   7620
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   8070
      X2              =   8310
      Y1              =   7620
      Y2              =   7920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   13050
      TabIndex        =   16
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LabelPacchi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pacchi"
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
      Height          =   405
      Left            =   6960
      TabIndex        =   15
      Top             =   6510
      Width           =   6225
   End
   Begin VB.Label DisplayDistGruppi 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "300"
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
      Left            =   12090
      TabIndex        =   14
      Top             =   7080
      Width           =   915
   End
   Begin VB.Label DisplayDistPacchi 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   10350
      TabIndex        =   13
      Top             =   7080
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   11100
      TabIndex        =   12
      Top             =   7050
      Width           =   315
   End
   Begin VB.Image Image6 
      Height          =   2835
      Left            =   480
      Picture         =   "FormStoccaggio.frx":2CA3
      Top             =   7440
      Width           =   14925
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Index           =   11
      Left            =   6780
      Top             =   6420
      Width           =   6645
   End
End
Attribute VB_Name = "frmStoccaggio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OrdineCorrente As OrderClass
    
'********************************************************
'       Variabili per pesa
'********************************************************
Private SOH As String
Private STX As String
Private CR As String
Private NUL As String
Private ETX As String
Private EOT As String
Private ENQ As String
Private ACK As String
Private POL As String

' **** timeout *******
Private Const Timeout As Integer = 1

Private WeightState As TWeightState
Private WeightDelayCounter As Integer
Private ActualWeight As Double
Private Const PesoNonValido As Integer = -9999
Dim WeightRequestString() As Byte
Dim WeightResetString() As Byte
Dim WeightErrorString() As Byte
Dim WeightReceivedData() As Byte
Private Const InBufferSize As Integer = 64
Private Enum TWeightState
    WInit = 10
    WStandby = 11
    WGrossRead = 31
    WErrorRead = 33
    WResetRead = 35
    WError = 40
    WPause = 41
End Enum

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub Image1_Click()
  DB486.Bit(62, 0) = Not DB486.Bit(62, 0)
End Sub

Private Sub Image5_Click()
  DB486.Bit(62, 1) = Not DB486.Bit(62, 1)
End Sub

Private Sub Label3_Click()
Dim temp As Double
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 1200
    TOUCHNumericPad.Dati = DB486.Word(72) / 10
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        DB486.Word(72) = temp * 10
            If (temp * 10) < DB486.Word(64) Then
               DB486.Word(64) = temp * 10
            End If
    End If
    RefreshDati
End Sub

Private Sub Label6_Click()
   Dim temp As Double
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 1200
    TOUCHNumericPad.Dati = DB486.Word(64) / 10
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        DB486.Word(64) = temp * 10
            If (temp * 10) > DB486.Word(72) Then
               DB486.Word(72) = temp * 10
            End If
    End If
    RefreshDati
End Sub

Private Sub SelScoli_Click()
    DB486.Bit(62, 0) = Not DB486.Bit(62, 0)
End Sub

Private Sub TimerLocale_Timer()
   Me.Update
End Sub
Public Sub Update()
   
   lblbar(2) = PaginaStoccaggio.Ordine_Descrizione
   lblbar(4) = PaginaStoccaggio.Ricetta_Descrizione
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
   If DB402.DatiCambiati Then
       RefreshDati
       DB402.DatiCambiati = False
   End If
   FramePrinter.Visible = PrinterInstall Or Param.GetBit("Par210_Simulazione_PLCOff")
   FrameStampa.Visible = Not Param.GetBit("Par208_PresenzaStampante")
End Sub
'Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
'   frmKernel.PaginaCorrente = Index
'End Sub

Private Sub Form_Activate()
    
    Me.Update
    TimerLocale.Enabled = True
    TimerLocale.Interval = 250
    ' (disattivare in Form_deactivate)
    Barra21.Selezionato = 6
    ' aggiornamento dati ordine attuale
    Set OrdineCorrente = New OrderClass
    OrdineCorrente.IDOrdine = frmKernel.CodOrdineCorrente.CodStoccaggio
     If Param.GetBit("Par204_AttivaPesa") Then
  '      FramePesa.Visible = True
        Shape2(7).Visible = True
     Else
   '     FramePesa.Visible = False
        Shape2(7).Visible = False
     End If
     FrameSoloScoli.Visible = Param.GetBit("Par225_Scoli_on_off")
     Label5.Visible = Param.GetBit("Par226_Portatavolette")
     Label01(4).Visible = Param.GetBit("Par226_Portatavolette")
     Image5.Visible = Param.GetBit("Par226_Portatavolette")
  '   FramePacchi.Visible = Param.GetBit("Par227_DistanziamentoPacchi")
     FrameStampa.Visible = Not Param.GetBit("Par208_PresenzaStampante")
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   TimerLocale.Enabled = True
   TimerLocale.Interval = 500
   
   'Picture = LoadPicture("..\bitmap\SfondoDP6_ver2_0.jpg")
'   For i = 0 To 7
'      ImgPaccoEsagono(i).Picture = LoadPicture("..\bitmap\paccoesagono.gif")
'   Next
   WindowState = 2
    ScritteMultilingua
    
    ' inizializza variabili
    ' ************ pesa ****************
    WeightState = WInit
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
    RefreshSelettoreStampa
End Sub

Private Sub RefreshDati()
    
    'selettore scoli
    If DB486.Bit(62, 0) Then
       SelScoli.Picture = ImageList1.ListImages(2).Picture
       Image1.Picture = ImageList1.ListImages(2).Picture
    Else
       SelScoli.Picture = ImageList1.ListImages(1).Picture
        Image1.Picture = ImageList1.ListImages(1).Picture
    End If
    ' portatavolette
     If DB486.Bit(62, 1) Then
       Image5.Picture = ImageList1.ListImages(2).Picture
    Else
        Image5.Picture = ImageList1.ListImages(1).Picture
    End If
    LabelCartellino(0).Visible = DB402.Bit(8, 2)
    WeightDisplay.Visible = LabelCartellino(0).Visible
    Shape2(2).Visible = WeightDisplay.Visible
  '  Shape12.Visible = LabelCartellino(0).Visible
    WeightResetCommand.Visible = False ' LabelCartellino(0).Visible
    'visualizzazione del peso
    If Param.GetBit("Par222_LetturaPeso") = False Then
            If LabelCartellino(0).Visible Then
                WeightSelector.Picture = ImageList1.ListImages(2).Picture
                If DB402.Word(16) <> PesoNonValido Then
                   If DB402.Word(16) > Param.GetNumber("DB402_014_BandaPesaAZero") Then
                       WeightDisplay.caption = Unit.kg_To_Display_kg(DB402.Word(16))
                    Else
                       WeightDisplay.caption = "0"
                    End If
                Else
                    WeightDisplay.caption = "- - -"
                End If
            Else
                WeightSelector.Picture = ImageList1.ListImages(1).Picture
                WeightDisplay.caption = "- - -"
            End If
    Else
        ' se modo profibus attivato visualizza sempre il peso
           Select Case DB402.Word(38) 'gestione errori
           Case 0
                 If DB402.Word(16) > Param.GetNumber("DB402_014_BandaPesaAZero") Then
                     WeightDisplay.caption = Unit.kg_To_Display_kg(DB402.Word(16))
                 Else
                    WeightDisplay.caption = "0"
                 End If
           Case Else
                    WeightDisplay.caption = "Err: " & DB402.Word(38)
           End Select
           If DB402.Bit(8, 2) Then
                WeightSelector.Picture = ImageList1.ListImages(2).Picture
            Else
                WeightSelector.Picture = ImageList1.ListImages(1).Picture
            End If
    End If
    
    Label6.caption = DB486.Word(64) / 10#
    ' Label11.caption = Label6.caption
    Label3.caption = DB486.Word(72) / 10#
    
    DisplayPacchi.caption = DB486.Word(66)
    DisplayDistGruppi.caption = DB486.Word(70) / 10#
    DisplayDistPacchi.caption = DB486.Word(68) / 10#
   
    ' dati ultimo pacco pesato
    DisplayCartellino(1).caption = DB402.Word(12)   ' numero pacco
    DisplayCartellino(2).caption = DB402.Word(14)   ' numero tubi
    'DisplayCartellino(3).Caption = DB486.Word(24)   ' peso pacco
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

Private Sub WeightResetCommand_Click()
    If Param.GetBit("Par222_LetturaPeso") Then
       DB402.Bit(8, 6) = 1
    Else
       DB402.Bit(8, 3) = 1
    End If
End Sub

' pulsante abilitazione pesa
Private Sub WeightSelector_Click()
    ' aggiorna flag di pesa abilitata
    DB402.Bit(8, 2) = Not DB402.Bit(8, 2)
'    ' fa ripartire il dfa di comunicazione
'    WeightState = WInit
End Sub

Private Sub PrinterSelector_Click()
    If Param.GetBit("Par112_StampaInserita") Then
        Param.SetBit "Par112_StampaInserita", False
    Else
        Param.SetBit "Par112_StampaInserita", True
    End If
    RefreshSelettoreStampa
End Sub

Private Sub RefreshSelettoreStampa()
    If Param.GetBit("Par112_StampaInserita") Then
        PrinterSelector.Picture = ImageList1.ListImages(2).Picture
        CommandModifica.Visible = True
        CommandPrint.Visible = True
    Else
        PrinterSelector.Picture = ImageList1.ListImages(1).Picture
        CommandModifica.Visible = False
        CommandPrint.Visible = False
    End If
End Sub

'********************************************************
' Funzioni di interfaccia pubblica per pesa e stampante
'********************************************************

' stato pesa
Public Function ValidWeight() As Boolean
    If DB402.Word(16) <> PesoNonValido Then
        ValidWeight = True
    Else
        ValidWeight = False
    End If
    If Not DB402.Bit(8, 2) Then ValidWeight = True
End Function

' modifica campi cartellino
Private Sub CommandModifica_Click()
    Modifica
End Sub
' modifica campi cartellino
Public Sub Modifica()
    RecipeModifyForm.SSTab1.TabEnabled(0) = False
    RecipeModifyForm.SSTab1.TabEnabled(4) = True
    RecipeModifyForm.SSTab2.Tab = 0
    RecipeModifyForm.SSTab1.Tab = 4
    RecipeModifyForm.SSTab2.TabEnabled(1) = True
    ' carica ordine per dati cartellino
'    ModOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodRegge
    ModOrder.UploadData DB402.Word(18)
    ModRecipe.UploadData ModOrder.IDRicetta
    ' aggiorna pagina modifica
   ' frmStampa.VariableLabelUpdate
    RecipeModifyForm.Update
    RecipeModifyForm.StrapEnable = False
    RecipeModifyForm.BundleEnable = False
    
    WeightToPrint = ModRecipe.Weight
     ' esegue la finestra modifica dati
    RecipeModifyForm.Show (vbModal)
    If RecipeModifyForm.OK Then
       ModOrder.DownloadData DB402.Word(18)
       ModRecipe.DownloadData ModOrder.IDRicetta
    End If
End Sub

' comando stampa etichetta
Public Sub CommandPrint_Click()
    With frmStampa
         '.LoadTicketFile "..\Target\Tickets\MairTicket_" & Format(Cartellino.Lingua, "00") & ".TKT"
''         .LoadFixTexts
''         .LengFixTextRefresh (Cartellino.Lingua)
''         .FixVisibleRefresh
''         .RefreshVar
         .PrintExec DB402.Word(18)
    End With
End Sub

Private Sub Timer1_Timer()

    If Param.GetBit("Par222_LetturaPeso") = False Then
        '*************************************************************
        '  Dfa pesa NOBEL
        '*************************************************************
        Select Case WeightState
            Case WInit
                WeightState = WStandby
                If DB402.Bit(8, 2) Then
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
                        If Err.number <> 0 Then MsgBox "Parameter ""Par115_SerialePesa"" error", vbOKOnly
                    On Error GoTo 0
                    On Error Resume Next
                        ' riapre porta
                        If Not WeightMSComm.PortOpen Then WeightMSComm.PortOpen = True
                        ' se non riesce ad aprirla significa che la porta non è disponibile
                        If Err.number <> 0 Then
                            frmKernel.PcAlarm(WeightCommPort) = True
                        Else
                            frmKernel.PcAlarm(WeightCommPort) = False
                            WeightMSComm.Output = WeightErrorString
                            WeightState = WErrorRead
                        End If
                    On Error GoTo 0
                Else
                    On Error Resume Next
                    If WeightMSComm.PortOpen Then WeightMSComm.PortOpen = False
                    frmKernel.PcAlarm(WeightCommPort) = False
                    frmKernel.PcAlarm(WeightCommErr) = False
                    frmKernel.PcAlarm(WeightFault) = False
                End If
                DB402.Word(16) = PesoNonValido
            Case WStandby   ' do nothing state if Weight is not enabled
                If DB402.Bit(8, 2) Then WeightState = WInit
            Case WErrorRead
                If WeightMSComm.InBufferCount >= 7 Then
                    WeightReceivedData = WeightMSComm.Input
                    If WeightReceivedData(4) <> 0 Then
                        DB402.Word(16) = PesoNonValido
                        frmKernel.PcAlarm(WeightFault) = True
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
                    ' controllo per evitare overflow
                    If ActualWeight > 32000 Then ActualWeight = 32000
                    If ActualWeight < -32000 Then ActualWeight = -32000
                    DB402.Word(16) = ActualWeight
    
                    frmKernel.PcAlarm(WeightFault) = False
                    frmKernel.PcAlarm(WeightCommErr) = False
                    WeightDelayCounter = 0
                    ' esecuzione comando azzeramento
                    If DB402.Bit(8, 3) Then
                        DB402.Bit(8, 3) = False
                        WeightMSComm.Output = WeightResetString
                        WeightState = WResetRead
'                    Else
'                        WeightMSComm.Output = WeightErrorString
'                        WeightState = WErrorRead
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
                'DB402.Word(16) = PesoNonValido  ' valore che indica mancanza di connessione alla pesa
                frmKernel.PcAlarm(WeightCommErr) = True
                frmKernel.PcAlarm(WeightFault) = False
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
    End If
End Sub


Private Sub DisplayPacchi_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 5
    TOUCHNumericPad.Dati = DB486.Word(66)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 5# Then temp = 5#
        If temp < 0 Then temp = 0
        DB486.Word(66) = temp
    End If
    RefreshDati
End Sub

Private Sub DisplayDistGruppi_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = DB486.Word(70) / 10#
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 100# Then temp = 100#
        If temp < 1# Then temp = 1#
        DB486.Word(70) = temp * 10#
    End If
    RefreshDati
End Sub

Private Sub DisplayDistPacchi_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 1
    TOUCHNumericPad.ValoreMin = 0
    TOUCHNumericPad.ValoreMax = 50
    TOUCHNumericPad.Dati = DB486.Word(68) / 10#
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 50# Then temp = 50#
        If temp < 0# Then temp = 0#
        DB486.Word(68) = temp * 10#
    End If
    RefreshDati
End Sub

Sub ScritteMultilingua()
    LabelPesa.caption = Param.Text("Pesa")
    LabelCartellino(0).caption = Param.Text("Pesoreale")
    LabelCartellino(2).caption = Param.Text("Numero pacco")
    LabelCartellino(3).caption = Param.Text("Numero barre")
    Label2(2).caption = Param.Text("Printer")
    lblbar(5) = Param.Text("Storage page")
    Label4.caption = Param.Text("Scoli")
    Label9.caption = Label4.caption
   ' Label7.caption = Param.Text("RitardoPesa")
    Label10.caption = Label7.caption
    CommandModifica.caption = Param.Text("F9 - Modify")
    CommandPrint.caption = Param.Text("F12 - Print")
    LabelPacchi.caption = Param.Text("Pacchi")
    lblbar(1) = Param.Text("Pagina")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
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
                .NomeFile = "Stoccaggio_pagina.htm"
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

