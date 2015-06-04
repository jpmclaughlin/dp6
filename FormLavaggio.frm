VERSION 5.00
Begin VB.Form FormLavaggio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillColor       =   &H00FFFF00&
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Height          =   2715
      Left            =   7620
      TabIndex        =   36
      Top             =   4950
      Width           =   3555
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   810
         Width           =   3285
      End
      Begin VB.Image Image5 
         Height          =   1335
         Left            =   780
         Picture         =   "FormLavaggio.frx":0000
         Top             =   1320
         Width           =   1965
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nozzle choose"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   2
         Left            =   150
         TabIndex        =   38
         Top             =   180
         Width           =   3270
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   5385
      Left            =   11160
      TabIndex        =   28
      Top             =   4950
      Width           =   4095
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nozzle diameter [mm]"
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
         Index           =   1
         Left            =   420
         TabIndex        =   35
         Top             =   2760
         Width           =   3210
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1260
         TabIndex        =   34
         Top             =   3240
         Width           =   1425
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual valve status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   3
         Left            =   750
         TabIndex        =   33
         Top             =   210
         Width           =   2880
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   570
         TabIndex        =   32
         Top             =   2010
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2550
         TabIndex        =   31
         Top             =   2010
         Width           =   705
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   4
         Left            =   840
         TabIndex        =   30
         Top             =   1470
         Width           =   180
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   600
         Index           =   5
         Left            =   2790
         TabIndex        =   29
         Top             =   1470
         Width           =   300
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   0
         Left            =   180
         Top             =   180
         Width           =   3795
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   660
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   585
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   525
         Left            =   2610
         Shape           =   3  'Circle
         Top             =   1440
         Width           =   585
      End
      Begin VB.Image Image4 
         Height          =   1320
         Index           =   0
         Left            =   270
         Picture         =   "FormLavaggio.frx":0220
         Top             =   540
         Width           =   1620
      End
      Begin VB.Image Image4 
         Height          =   1320
         Index           =   1
         Left            =   2190
         Picture         =   "FormLavaggio.frx":06A6
         Top             =   540
         Width           =   1620
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2715
      Left            =   30
      TabIndex        =   20
      Top             =   4950
      Width           =   4185
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
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
         Height          =   315
         Index           =   8
         Left            =   3360
         TabIndex        =   27
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label DisplayDiametroBarra 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1890
         TabIndex        =   26
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label LabelDiametroBarra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Diameter"
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
         Left            =   -360
         TabIndex        =   25
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
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
         Height          =   315
         Index           =   9
         Left            =   3360
         TabIndex        =   24
         Top             =   1710
         Width           =   795
      End
      Begin VB.Label LabelLungezzaBarra 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lenght"
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
         Left            =   -360
         TabIndex        =   23
         Top             =   1710
         Width           =   1935
      End
      Begin VB.Label DisplayLunghezzaBarra 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1860
         TabIndex        =   22
         Top             =   1650
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tube"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Index           =   9
         Left            =   630
         TabIndex        =   21
         Top             =   180
         Width           =   3210
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2715
      Left            =   4230
      TabIndex        =   16
      Top             =   4950
      Width           =   3375
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   750
         Width           =   3210
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   960
         TabIndex        =   18
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Scharf breacker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Index           =   8
         Left            =   90
         TabIndex        =   17
         Top             =   180
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   52
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1050
         Left            =   3180
         Picture         =   "FormLavaggio.frx":0B2C
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
         TabIndex        =   13
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   510
      Top             =   360
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   14
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin dp6.ControlloUpDown ControlloTraspLav 
      Height          =   990
      Left            =   180
      TabIndex        =   50
      Top             =   3540
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1746
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Conveyor speed"
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
      Height          =   510
      Left            =   180
      TabIndex        =   51
      Top             =   3000
      Width           =   3090
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   2
      Left            =   150
      Top             =   2940
      Width           =   3285
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   13
      Left            =   8580
      TabIndex        =   49
      Top             =   2610
      Width           =   225
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   12
      Left            =   7920
      TabIndex        =   48
      Top             =   2580
      Width           =   225
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   11
      Left            =   6840
      TabIndex        =   47
      Top             =   2610
      Width           =   225
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Index           =   10
      Left            =   6240
      TabIndex        =   46
      Top             =   2580
      Width           =   225
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   7860
      X2              =   7860
      Y1              =   3120
      Y2              =   2550
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   555
      Index           =   4
      Left            =   11460
      TabIndex        =   45
      Top             =   4380
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   555
      Index           =   3
      Left            =   9450
      TabIndex        =   44
      Top             =   1860
      Width           =   1425
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7860
      X2              =   10620
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   8520
      X2              =   8520
      Y1              =   3120
      Y2              =   2550
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   7170
      X2              =   7170
      Y1              =   2550
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6510
      X2              =   6510
      Y1              =   2550
      Y2              =   3120
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4620
      X2              =   7170
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   555
      Index           =   5
      Left            =   11640
      TabIndex        =   43
      Top             =   4290
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   555
      Index           =   2
      Left            =   4530
      TabIndex        =   42
      Top             =   1860
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blowing time [s]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   7
      Left            =   7365
      TabIndex        =   41
      Top             =   1950
      Width           =   1980
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blowing time [s]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   6
      Left            =   2415
      TabIndex        =   40
      Top             =   1920
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1          2         3          4         5         6         7        8         9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4980
      TabIndex        =   39
      Top             =   3630
      Width           =   5595
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00EE3959&
      BackStyle       =   0  'Transparent
      Caption         =   "Low speed"
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
      Height          =   285
      Left            =   11460
      TabIndex        =   2
      Top             =   3930
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00EE3959&
      BackStyle       =   0  'Transparent
      Caption         =   "Low speed"
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
      Height          =   285
      Left            =   11460
      TabIndex        =   1
      Top             =   3300
      Width           =   1260
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00EE3959&
      BackStyle       =   0  'Transparent
      Caption         =   "Low speed"
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
      Height          =   285
      Left            =   11490
      TabIndex        =   0
      Top             =   2670
      Width           =   1260
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   50
      Left            =   14670
      Picture         =   "FormLavaggio.frx":2BBA
      Top             =   3840
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   49
      Left            =   14670
      Picture         =   "FormLavaggio.frx":2F77
      Top             =   3210
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   48
      Left            =   14670
      Picture         =   "FormLavaggio.frx":3334
      Top             =   2550
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   2055
      Index           =   10
      Left            =   11400
      Top             =   2400
      Width           =   3825
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Legend"
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
      Height          =   315
      Left            =   11820
      TabIndex        =   15
      Top             =   2010
      Width           =   3105
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   0
      Left            =   4920
      Picture         =   "FormLavaggio.frx":36F1
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   1
      Left            =   5550
      Picture         =   "FormLavaggio.frx":3AAE
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   2
      Left            =   6180
      Picture         =   "FormLavaggio.frx":3E6B
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   3
      Left            =   6930
      Picture         =   "FormLavaggio.frx":4228
      Top             =   3090
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   4
      Left            =   7650
      Picture         =   "FormLavaggio.frx":45E5
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   5
      Left            =   8310
      Picture         =   "FormLavaggio.frx":49A2
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   6
      Left            =   8970
      Picture         =   "FormLavaggio.frx":4D5F
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   7
      Left            =   9630
      Picture         =   "FormLavaggio.frx":511C
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   24
      Left            =   9630
      Picture         =   "FormLavaggio.frx":54D9
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   25
      Left            =   8970
      Picture         =   "FormLavaggio.frx":5896
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   26
      Left            =   8310
      Picture         =   "FormLavaggio.frx":5C53
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   27
      Left            =   7650
      Picture         =   "FormLavaggio.frx":6010
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   28
      Left            =   6930
      Picture         =   "FormLavaggio.frx":63CD
      Top             =   3090
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   29
      Left            =   6180
      Picture         =   "FormLavaggio.frx":678A
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   30
      Left            =   5550
      Picture         =   "FormLavaggio.frx":6B47
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   31
      Left            =   4920
      Picture         =   "FormLavaggio.frx":6F04
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   40
      Left            =   9630
      Picture         =   "FormLavaggio.frx":72C1
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   41
      Left            =   8970
      Picture         =   "FormLavaggio.frx":767E
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   42
      Left            =   8310
      Picture         =   "FormLavaggio.frx":7A3B
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   43
      Left            =   7650
      Picture         =   "FormLavaggio.frx":7DF8
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   44
      Left            =   6930
      Picture         =   "FormLavaggio.frx":81B5
      Top             =   3090
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   45
      Left            =   6180
      Picture         =   "FormLavaggio.frx":8572
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   46
      Left            =   5550
      Picture         =   "FormLavaggio.frx":892F
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   47
      Left            =   4920
      Picture         =   "FormLavaggio.frx":8CEC
      Top             =   3060
      Width           =   450
   End
   Begin VB.Image Image3 
      Height          =   1050
      Index           =   0
      Left            =   3330
      Picture         =   "FormLavaggio.frx":90A9
      Stretch         =   -1  'True
      Top             =   2670
      Width           =   1440
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   8
      Left            =   10290
      Picture         =   "FormLavaggio.frx":9435
      Top             =   3060
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   9
      Left            =   12630
      Picture         =   "FormLavaggio.frx":97F2
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   10
      Left            =   13290
      Picture         =   "FormLavaggio.frx":9BAF
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   11
      Left            =   11940
      Picture         =   "FormLavaggio.frx":9F6C
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   12
      Left            =   12630
      Picture         =   "FormLavaggio.frx":A329
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   13
      Left            =   13320
      Picture         =   "FormLavaggio.frx":A6E6
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   14
      Left            =   13950
      Picture         =   "FormLavaggio.frx":AAA3
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   15
      Left            =   12930
      Picture         =   "FormLavaggio.frx":AE60
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   16
      Left            =   13230
      Picture         =   "FormLavaggio.frx":B21D
      Top             =   3780
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   17
      Left            =   13950
      Picture         =   "FormLavaggio.frx":B5DA
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   18
      Left            =   13320
      Picture         =   "FormLavaggio.frx":B997
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   19
      Left            =   12630
      Picture         =   "FormLavaggio.frx":BD54
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   20
      Left            =   11940
      Picture         =   "FormLavaggio.frx":C111
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   21
      Left            =   13290
      Picture         =   "FormLavaggio.frx":C4CE
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   22
      Left            =   12630
      Picture         =   "FormLavaggio.frx":C88B
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   23
      Left            =   10290
      Picture         =   "FormLavaggio.frx":CC48
      Top             =   3060
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   32
      Left            =   13140
      Picture         =   "FormLavaggio.frx":D005
      Top             =   3750
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   33
      Left            =   13950
      Picture         =   "FormLavaggio.frx":D3C2
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   34
      Left            =   13320
      Picture         =   "FormLavaggio.frx":D77F
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   35
      Left            =   12630
      Picture         =   "FormLavaggio.frx":DB3C
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   36
      Left            =   11940
      Picture         =   "FormLavaggio.frx":DEF9
      Top             =   3870
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   37
      Left            =   13290
      Picture         =   "FormLavaggio.frx":E2B6
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   38
      Left            =   12630
      Picture         =   "FormLavaggio.frx":E673
      Top             =   3360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   39
      Left            =   10290
      Picture         =   "FormLavaggio.frx":EA30
      Top             =   3060
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   3
      Left            =   11400
      Top             =   1980
      Width           =   3825
   End
   Begin VB.Image ImgWB 
      Height          =   990
      Left            =   4380
      Picture         =   "FormLavaggio.frx":EDED
      Top             =   3150
      Width           =   6675
   End
End
Attribute VB_Name = "FormLavaggio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Scritta(15) As New cScrittaVert
Dim i

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
      frmKernel.PaginaCorrente = Index
End Sub

Private Sub Combo1_Click()
   DB465.Word(20) = Combo1.ListIndex
End Sub





'Private Sub DisplayDiametroBarra_Click()
'        TOUCHNumericPad.Decimali = 1
'        TOUCHNumericPad.ValoreMin = Unit.m_To_Display_mm(Param.GetNumber("Par002_Tubo_LarghezzaMin"))
'        TOUCHNumericPad.ValoreMax = Unit.m_To_Display_mm(Param.GetNumber("Par001_Tubo_LarghezzaMax"))
'        TOUCHNumericPad.Dati = DisplayDiametroBarra
'        TOUCHNumericPad.Show vbModal
'        If TOUCHNumericPad.DatiConfermati Then
'            Label5(2).Caption = TOUCHNumericPad.Dati
'            DB465.Word(6) = TOUCHNumericPad.Dati * 10
'        End If
'
'End Sub

'Private Sub DisplayLunghezzaBarra_Click()
'        TOUCHNumericPad.Decimali = 0
'        TOUCHNumericPad.ValoreMin = Unit.m_To_Display_mm(Param.GetNumber("Par008_Tubo_LunghezzaMin"))
'        TOUCHNumericPad.ValoreMax = Unit.m_To_Display_mm(Param.GetNumber("Par007_Tubo_LunghezzaMax"))
'        TOUCHNumericPad.Dati = DisplayLunghezzaBarra
'        TOUCHNumericPad.Show vbModal
'        If TOUCHNumericPad.DatiConfermati Then
'            Label5(2).Caption = TOUCHNumericPad.Dati
'            DB465.Word(4) = TOUCHNumericPad.Dati
'        End If
'End Sub

Private Sub Label5_Click(Index As Integer)
    If Index = 2 Then
        TOUCHNumericPad.Decimali = 0
         TOUCHNumericPad.ValoreMin = 0
        TOUCHNumericPad.ValoreMax = 10
        TOUCHNumericPad.Dati = Label5(2)
        TOUCHNumericPad.Show vbModal
        If TOUCHNumericPad.DatiConfermati Then
            Label5(2).Caption = TOUCHNumericPad.Dati
            DB465.Word(22) = TOUCHNumericPad.Dati * 10
        End If
    End If
    If Index = 5 Then
        TOUCHNumericPad.Decimali = 0
        TOUCHNumericPad.ValoreMin = 0
        TOUCHNumericPad.ValoreMax = 10
        TOUCHNumericPad.Dati = Label5(5)
        TOUCHNumericPad.Show vbModal
        If TOUCHNumericPad.DatiConfermati Then
            Label5(5).Caption = TOUCHNumericPad.Dati
            DB465.Word(24) = TOUCHNumericPad.Dati * 10
        End If
    End If
    If Index = 4 Then
        TOUCHNumericPad.Decimali = 0
        TOUCHNumericPad.ValoreMin = 0
        TOUCHNumericPad.ValoreMax = 10
        TOUCHNumericPad.Dati = Label5(4)
        TOUCHNumericPad.Show vbModal
        If TOUCHNumericPad.DatiConfermati Then
            Label5(4).Caption = TOUCHNumericPad.Dati
            DB465.Word(26) = TOUCHNumericPad.Dati * 10
        End If
    End If
    If Index = 3 Then
        TOUCHNumericPad.Decimali = 0
        TOUCHNumericPad.ValoreMin = 0
        TOUCHNumericPad.ValoreMax = 10
        TOUCHNumericPad.Dati = Label5(3)
        TOUCHNumericPad.Show vbModal
        If TOUCHNumericPad.DatiConfermati Then
            Label5(3).Caption = TOUCHNumericPad.Dati
            DB465.Word(28) = TOUCHNumericPad.Dati * 10
        End If
    End If
    
End Sub

Private Sub TimerLocale_Timer()
      Me.Update
End Sub

Public Sub Update()
    Dim i As Integer
    Dim mask As Integer

    Label5(2) = DB465.Word(22) / 10
    Label5(5) = DB465.Word(24) / 10
    Label5(4) = DB465.Word(26) / 10
    Label5(3) = DB465.Word(28) / 10
    lblbar(2) = PaginaWb.Ordine_Descrizione
    lblbar(4) = PaginaWb.Ricetta_Descrizione
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
     If DB418.MaskBit(56, 0) = False Then
        Label5(6) = "OFF"
        Label5(6).BackColor = vbGreen
     Else
        Label5(6) = "ON"
        Label5(6).BackColor = vbRed
     End If
     If DB418.MaskBit(56, 1) = False Then
        Label5(7) = "OFF"
        Label5(7).BackColor = vbGreen
     Else
        Label5(7) = "ON"
        Label5(7).BackColor = vbRed
     End If
     Label5(0) = DB418.Word(54)
     Label5(1) = DB418.Word(52)
     Dim a
     
    For i = 0 To 15
        mask = Abs(DB418.MaskBit(28, i)) Or 2 * Abs(DB418.MaskBit(30, i)) Or 4 * Abs(DB418.MaskBit(32, i))
        Select Case mask
        Case 5
            ImgTubo(15 - i).Visible = True
            ImgTubo(16 + i).Visible = False
            ImgTubo(32 + i).Visible = False
        Case 7
            ImgTubo(15 - i).Visible = False
            ImgTubo(16 + i).Visible = False
            ImgTubo(32 + i).Visible = True
        Case 4, 6
            ImgTubo(15 - i).Visible = False
            ImgTubo(16 + i).Visible = False
            ImgTubo(32 + i).Visible = False
'        Case 5, 7, 4
'            ImgTubo(15 - i).Visible = False
'            ImgTubo(16 + i).Visible = True
'            ImgTubo(32 + i).Visible = False
'        Case 7
'            ImgTubo(15 - i).Visible = False
'            ImgTubo(16 + i).Visible = False
'            ImgTubo(32 + i).Visible = True
        Case Else
            ImgTubo(15 - i).Visible = False
            ImgTubo(16 + i).Visible = True
            ImgTubo(32 + i).Visible = False
       End Select
     Next i
     
        'visualizzza la lunghezza della barra
        DisplayLunghezzaBarra.Caption = Unit.m_To_Display_mm(DB465.Word(4) / 1000)
        'visualizza le dimensioni del tubo
'        If DB465.Bit(2, 0) = True Then
'           DisplayAltezza.Visible = False
'           Label3(0).Visible = False
           DisplayDiametroBarra.Caption = Unit.m_To_Display_mm0(DB465.Word(6) / 10000#)
'        Else
'           DisplayAltezza.Visible = True
'           Label3(0).Visible = True
'           DisplayDiametroBarra.Caption = Unit.m_To_Display_mm0(DB465.Word(6) / 10000#)
'           DisplayAltezza.Caption = Unit.m_To_Display_mm0(DB460.Word(8) / 10000#)
'        End If
         If ControlloTraspLav.Occupato = False Then
           ControlloTraspLav.value = DB465.Word(32)
           ControlloTraspLav.Refresh
        End If
         

End Sub

Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

'Private Sub DisplayAltezza_Click()
'  Dim temp As Double
'    TOUCHNumericPad.Decimali = 1
'    TOUCHNumericPad.ValoreMin = Param.GetNumber("Par004_Tubo_AltezzaMin") * 1000
'    TOUCHNumericPad.ValoreMax = Param.GetNumber("Par003_Tubo_AltezzaMax") * 1000
'    TOUCHNumericPad.Dati = DB460.Word(8) / 10#
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati * 10
'        If temp > 5000# Then temp = 5000#
'        If temp < 10# Then temp = 10#
'        DB460.Word(8) = temp
'        DB460.MaskBit(20, 0) = True     ' dati modificati, una modifica del diametro va propagata al resto della linea
'    End If
'    Me.Update
'End Sub

'Private Sub DisplayDiametroBarra_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 1
'    TOUCHNumericPad.ValoreMin = Param.GetNumber("Par002_Tubo_LarghezzaMin") * 1000
'    TOUCHNumericPad.ValoreMax = Param.GetNumber("Par001_Tubo_LarghezzaMax") * 1000
'    TOUCHNumericPad.Dati = DB460.Word(6) / 10#
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati * 10
'        If temp > 5000# Then temp = 5000#
'        If temp < 10# Then temp = 10#
'        DB465.Word(6) = temp
'        DB465.MaskBit(20, 0) = True     ' dati modificati, una modifica del diametro va propagata al resto della linea
'    End If
'    Me.Update
'End Sub
'
'Private Sub DisplayLunghezzaBarra_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = Param.GetNumber("Par008_Tubo_LunghezzaMin") * 1000
'    TOUCHNumericPad.ValoreMax = Param.GetNumber("Par007_Tubo_LunghezzaMax") * 1000
'    TOUCHNumericPad.Dati = DB460.Word(4)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati
'        If temp > 32000# Then temp = 32000#
'        If temp < 500# Then temp = 500#
'        DB465.Word(4) = temp
'        DB465.MaskBit(20, 0) = True     ' dati modificati, una modifica della lunghezza va propagata al resto della linea
'    End If
'    Me.Update
'End Sub
Public Sub AggiornaDaUpDownControl()
    DB465.Word(32) = ControlloTraspLav.value
End Sub
Private Sub Form_Activate()
    TimerLocale.Enabled = True
    TimerLocale.Interval = 500

    Barra21.Selezionato = 14
    WindowState = vbMaximized
      
    ControlloTraspLav.Step = 5
    ControlloTraspLav.LimMax = 100
    ControlloTraspLav.LimMin = 10
    ControlloTraspLav.Decimali = 0
    ControlloTraspLav.Refresh
  
    Label5(1).Visible = Param.GetBit("Par181_Lavaggio")
    Label4(1).Visible = Label5(1).Visible
    Frame5.Visible = Label5(1).Visible
      ' abilitazione temporizzatore locale
    TimerLocale.Enabled = True

    Combo1.Clear
    Combo1.AddItem "Nothing"
    Combo1.AddItem "Blowing"
    Combo1.AddItem "Double blow."
    If Param.GetBit("Par181_Lavaggio") Then Combo1.AddItem "Washing"
    Combo1.ListIndex = Abs(DB465.Word(20))
    Me.Update
    
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
End Sub

Private Sub Form_Resize()
   For i = 0 To 15
      Scritta(i).Draw
   Next
End Sub

'Private Sub VelocLentaDisplay_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = 20
'    TOUCHNumericPad.ValoreMax = 100
'    TOUCHNumericPad.Dati = DB465.Word(20)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati
'        If temp > 100# Then temp = 100#
'        If temp < 20 Then temp = 20#
'        DB465.Word(20) = temp
'    End If
'    Me.Update
'End Sub
''
'Private Sub VelocRapidaDisplay_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = 20
'    TOUCHNumericPad.ValoreMax = 100
'    TOUCHNumericPad.Dati = DB465.Word(22)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati
'        If temp > 100# Then temp = 100#
'        If temp < 20 Then temp = 20#
'        DB465.Word(22) = temp
'    End If
'    Me.Update
'End Sub
'
'Private Sub VelocWBDisplay_Click()
'    Dim temp As Double
'    TOUCHNumericPad.Decimali = 0
'    TOUCHNumericPad.ValoreMin = 20
'    TOUCHNumericPad.ValoreMax = 100
'    TOUCHNumericPad.Dati = DB465.Word(18)
'    TOUCHNumericPad.Show vbModal
'    If TOUCHNumericPad.DatiConfermati Then
'        temp = TOUCHNumericPad.Dati
'        If temp > 100# Then temp = 100#
'        If temp < 20 Then temp = 20#
'        DB465.Word(18) = temp
'    End If
'    Me.Update
'End Sub


Private Sub Form_Load()
   
    ScritteMultilingua
    TimerLocale.Enabled = False
'   For i = 0 To 15
'     ' picLogo(i).Visible = False
'      Scritta(i).DrawingObject = picLogo(i)
'      Scritta(i).StartColor = &HFFFF00   '&HC0C000   '&HFF00FF
'      Scritta(i).EndColor = &HFFFFFF
'   Next
        
'      Scritta(5).StartColor = &H808080
'      Scritta(5).EndColor = &HFFFFFF
'      Scritta(3).StartColor = &H808080
'      Scritta(3).EndColor = &HFFFFFF
'      Scritta(10).StartColor = &H808080
'      Scritta(10).EndColor = &HFFFFFF
'      Scritta(12).StartColor = &H808080
'      Scritta(12).EndColor = &HFFFFFF
'      Scritta(8).StartColor = &HFFFF&
'      Scritta(8).EndColor = &HFFFFFF
'      Scritta(7).StartColor = &HFFFF&
'      Scritta(7).EndColor = &HFFFFFF
'      Scritta(14).StartColor = &HFFFF&
'      Scritta(14).EndColor = &HFFFFFF
   
   'posizione data e ora
 
   For i = 0 To 47
      If i < 16 Then ImgTubo(i).Picture = LoadPicture("..\bitmap\tubo.gif")
      If i > 15 And i < 32 Then ImgTubo(i).Picture = LoadPicture("..\bitmap\TuboRosso.gif")
      If i > 31 Then ImgTubo(i).Picture = LoadPicture("..\bitmap\tuboverde.gif")
   Next
   
   ImgTubo(48).Picture = LoadPicture("..\bitmap\tubo.gif")
   ImgTubo(49).Picture = LoadPicture("..\bitmap\tuboverde.gif")
   ImgTubo(50).Picture = LoadPicture("..\bitmap\tuborosso.gif")
   ImgWB.Picture = LoadPicture("..\bitmap\trasplavaggio.gif")
   WindowState = 2
End Sub

Sub ScritteMultilingua()
     LabelDiametroBarra.Caption = Param.Text("Diametro")
    LabelLungezzaBarra.Caption = Param.Text("Lunghezza")
 '   Label4.Caption = Param.Text("Low speed")
 '   Label6.Caption = Param.Text("Hi speed")
 '   Label7.Caption = Param.Text("Allineamento")

     '======================================================================
'    Scritta(15).RientroVert = 16
'    Scritta(15).Caption = Param.Text("Prelievo")
'    Scritta(14).RientroVert = 16
'    Scritta(14).Caption = Param.Text("Allineam. LOW")
'    Scritta(13).RientroVert = 16
'    Scritta(13).Caption = Param.Text("Filett.") & " 1"
'    Scritta(12).RientroVert = 16
'    Scritta(12).Caption = Param.Text("Sosta")
'    Scritta(11).RientroVert = 16
'    Scritta(11).Caption = Param.Text("Vernic.") & " 1"
'    Scritta(10).RientroVert = 16
'    Scritta(10).Caption = Param.Text("Sosta")
'    Scritta(9).RientroVert = 16
'    Scritta(9).Caption = Param.Text("Deposito")
'    Scritta(8).RientroVert = 16
'    Scritta(8).Caption = Param.Text("VR_Vel")
'    Scritta(7).RientroVert = 16
'    Scritta(7).Caption = "Soffiatura 1"  ' Param.Text("VR_Vel")
    Scritta(6).RientroVert = 16
    Scritta(6).Caption = "Soffiatura 2"  ' Param.Text("Smussat") & " 2"
  '  Scritta(5).RientroVert = 16
  '  Scritta(5).Caption = "Soffiatura 1" ' Param.Text("Allineam. LOW")
    Scritta(4).RientroVert = 16
    Scritta(4).Caption = "Soffiatura 1" 'Param.Text("VR_Vel")
    'Scritta(3).RientroVert = 16
    'Scritta(3).Caption = "Lavaggio 2" 'Param.Text("VR_Vel")
    Scritta(2).RientroVert = 16
    Scritta(2).Caption = "Lavaggio 2" 'Param.Text("Smussat") & " 1"
    Scritta(1).RientroVert = 16
    Scritta(1).Caption = "Lavaggio 1" 'Param.Text("Allineamento")
    Scritta(0).RientroVert = 16
    Scritta(0).Caption = Param.Text("Allineamento") 'Param.Text("Prelievo")
    '======================================================================
    Label8.Caption = Param.Text("Presenza")
    Label10.Caption = Param.Text("Completo")
    Label11.Caption = Param.Text("Non_abilitato")
    lblbar(5) = Param.Text("LavaggioPag")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
   ' Label8.Caption = Param.Text("LungTubo")
    lblbar(1) = Param.Text("Pagina")
   ' Label4 = Param.Text("Velmodifica")
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
                .NomeFile = "WB_pagina.htm"
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

