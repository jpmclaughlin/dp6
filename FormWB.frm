VERSION 5.00
Begin VB.Form FormWB 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   11400
      ScaleHeight     =   495
      ScaleWidth      =   915
      TabIndex        =   45
      Top             =   5550
      Width           =   945
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   3735
      Left            =   4590
      TabIndex        =   36
      Top             =   1170
      Width           =   5595
      Begin dp6.ControlloUpDown ControlloMonobeam3 
         Height          =   990
         Left            =   1200
         TabIndex        =   37
         Top             =   2430
         Visible         =   0   'False
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1746
      End
      Begin dp6.ControlloUpDown ControlloUpDownWB 
         Height          =   960
         Left            =   1230
         TabIndex        =   38
         Top             =   810
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1693
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Velocità via rulli (s)"
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
         Left            =   930
         TabIndex        =   40
         Top             =   1980
         Visible         =   0   'False
         Width           =   3930
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Velocità via rulli (s)"
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
         Height          =   360
         Left            =   630
         TabIndex        =   39
         Top             =   360
         Width           =   4590
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   1
         Left            =   960
         Top             =   270
         Width           =   3945
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   2
         Left            =   930
         Top             =   1920
         Visible         =   0   'False
         Width           =   3945
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   3735
      Left            =   210
      TabIndex        =   31
      Top             =   1170
      Width           =   4275
      Begin VB.Label Label7 
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
         Height          =   585
         Left            =   1050
         TabIndex        =   48
         Top             =   2940
         Width           =   1875
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wb position [°]"
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
         Height          =   615
         Left            =   1050
         TabIndex        =   47
         Top             =   2370
         Width           =   1875
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
         Height          =   495
         Left            =   1110
         TabIndex        =   35
         Top             =   630
         Width           =   1845
      End
      Begin VB.Label LabelDiametroBarra 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Diametro"
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
         Left            =   1170
         TabIndex        =   34
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label LabelLungezzaBarra 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Lunghezza"
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
         Left            =   1080
         TabIndex        =   33
         Top             =   1260
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
         Height          =   585
         Left            =   1080
         TabIndex        =   32
         Top             =   1680
         Width           =   1875
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   465
         Index           =   3
         Left            =   1140
         Top             =   180
         Width           =   1815
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   495
         Index           =   4
         Left            =   1110
         Top             =   1200
         Width           =   1845
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   0
      TabIndex        =   17
      Top             =   -60
      Width           =   15375
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   3
         Left            =   13800
         TabIndex        =   18
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Help"
         TxtTop          =   35
         TxtLeft         =   55
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   4
         ICONA           =   "..\Bitmap\manuale.ico"
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
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   2
         Left            =   12270
         TabIndex        =   19
         Top             =   150
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         TxtText         =   "Com"
         TxtTop          =   35
         TxtLeft         =   50
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
         TabIndex        =   20
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
         TabIndex        =   21
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   270
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   1050
         Left            =   3180
         Picture         =   "FormWB.frx":0000
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
         TabIndex        =   28
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   14
      Left            =   11850
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   16
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   0
      Left            =   2190
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   15
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   1
      Left            =   2880
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   14
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   2
      Left            =   3600
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   13
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   4
      Left            =   5010
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   12
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   6
      Left            =   6420
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   11
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   7
      Left            =   7140
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   10
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   8
      Left            =   7830
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   9
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   9
      Left            =   8490
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   8
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   11
      Left            =   9810
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   7
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   13
      Left            =   11160
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   15
      Left            =   12480
      ScaleHeight     =   2145
      ScaleWidth      =   285
      TabIndex        =   5
      Top             =   7710
      Width           =   315
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   7650
      ScaleHeight     =   495
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   5550
      Width           =   945
      Begin VB.Label VelocLentaDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   6600
      ScaleHeight     =   495
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   5550
      Width           =   945
      Begin VB.Label VelocRapidaDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   120
      Top             =   1020
   End
   Begin dp6.Barra2 Barra21 
      Height          =   1215
      Left            =   0
      TabIndex        =   30
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
   End
   Begin VB.Image imgDir 
      Height          =   480
      Index           =   1
      Left            =   11700
      Picture         =   "FormWB.frx":208E
      Top             =   6090
      Width           =   480
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   48
      Left            =   14610
      Picture         =   "FormWB.frx":24D0
      Top             =   1920
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   49
      Left            =   14610
      Picture         =   "FormWB.frx":288D
      Top             =   2460
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   50
      Left            =   14610
      Picture         =   "FormWB.frx":2C4A
      Top             =   3000
      Width           =   450
   End
   Begin VB.Label Label8 
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
      Height          =   525
      Left            =   11400
      TabIndex        =   44
      Top             =   1980
      Width           =   4095
   End
   Begin VB.Label Label10 
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
      Height          =   375
      Left            =   11400
      TabIndex        =   43
      Top             =   2550
      Width           =   4095
   End
   Begin VB.Label Label11 
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
      Height          =   405
      Left            =   11430
      TabIndex        =   42
      Top             =   3060
      Width           =   4095
   End
   Begin VB.Label Label5 
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
      Left            =   11790
      TabIndex        =   41
      Top             =   1470
      Width           =   3105
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modifica velocità"
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
      Height          =   390
      Left            =   5370
      TabIndex        =   29
      Top             =   5040
      Width           =   4710
   End
   Begin VB.Image imgDir 
      Height          =   480
      Index           =   0
      Left            =   7710
      Picture         =   "FormWB.frx":3007
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image imgDir 
      Height          =   480
      Index           =   4
      Left            =   7020
      Picture         =   "FormWB.frx":3449
      Top             =   6120
      Width           =   480
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   0
      Left            =   2130
      Picture         =   "FormWB.frx":388B
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   1
      Left            =   2790
      Picture         =   "FormWB.frx":3C48
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   2
      Left            =   3540
      Picture         =   "FormWB.frx":4005
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   3
      Left            =   4230
      Picture         =   "FormWB.frx":43C2
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   4
      Left            =   4920
      Picture         =   "FormWB.frx":477F
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   5
      Left            =   5640
      Picture         =   "FormWB.frx":4B3C
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   6
      Left            =   6330
      Picture         =   "FormWB.frx":4EF9
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   7
      Left            =   7050
      Picture         =   "FormWB.frx":52B6
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   8
      Left            =   7710
      Picture         =   "FormWB.frx":5673
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   9
      Left            =   8400
      Picture         =   "FormWB.frx":5A30
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   10
      Left            =   9060
      Picture         =   "FormWB.frx":5DED
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   11
      Left            =   9720
      Picture         =   "FormWB.frx":61AA
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   12
      Left            =   10410
      Picture         =   "FormWB.frx":6567
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   13
      Left            =   11100
      Picture         =   "FormWB.frx":6924
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   14
      Left            =   11730
      Picture         =   "FormWB.frx":6CE1
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   15
      Left            =   12420
      Picture         =   "FormWB.frx":709E
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   16
      Left            =   12420
      Picture         =   "FormWB.frx":745B
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   17
      Left            =   11730
      Picture         =   "FormWB.frx":7818
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   18
      Left            =   11100
      Picture         =   "FormWB.frx":7BD5
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   19
      Left            =   10410
      Picture         =   "FormWB.frx":7F92
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   20
      Left            =   9720
      Picture         =   "FormWB.frx":834F
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   21
      Left            =   9060
      Picture         =   "FormWB.frx":870C
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   22
      Left            =   8400
      Picture         =   "FormWB.frx":8AC9
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   23
      Left            =   7710
      Picture         =   "FormWB.frx":8E86
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   24
      Left            =   7050
      Picture         =   "FormWB.frx":9243
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   25
      Left            =   6330
      Picture         =   "FormWB.frx":9600
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   26
      Left            =   5640
      Picture         =   "FormWB.frx":99BD
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   27
      Left            =   4920
      Picture         =   "FormWB.frx":9D7A
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   28
      Left            =   4230
      Picture         =   "FormWB.frx":A137
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   29
      Left            =   3540
      Picture         =   "FormWB.frx":A4F4
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   30
      Left            =   2790
      Picture         =   "FormWB.frx":A8B1
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   31
      Left            =   2130
      Picture         =   "FormWB.frx":AC6E
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   32
      Left            =   12420
      Picture         =   "FormWB.frx":B02B
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   33
      Left            =   11730
      Picture         =   "FormWB.frx":B3E8
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   34
      Left            =   11100
      Picture         =   "FormWB.frx":B7A5
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   35
      Left            =   10410
      Picture         =   "FormWB.frx":BB62
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   36
      Left            =   9720
      Picture         =   "FormWB.frx":BF1F
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   37
      Left            =   9060
      Picture         =   "FormWB.frx":C2DC
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   38
      Left            =   8400
      Picture         =   "FormWB.frx":C699
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   39
      Left            =   7710
      Picture         =   "FormWB.frx":CA56
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   40
      Left            =   7050
      Picture         =   "FormWB.frx":CE13
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   41
      Left            =   6330
      Picture         =   "FormWB.frx":D1D0
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   42
      Left            =   5640
      Picture         =   "FormWB.frx":D58D
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   43
      Left            =   4920
      Picture         =   "FormWB.frx":D94A
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   44
      Left            =   4230
      Picture         =   "FormWB.frx":DD07
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   45
      Left            =   3540
      Picture         =   "FormWB.frx":E0C4
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   46
      Left            =   2790
      Picture         =   "FormWB.frx":E481
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image ImgTubo 
      Height          =   495
      Index           =   47
      Left            =   2130
      Picture         =   "FormWB.frx":E83E
      Top             =   6570
      Width           =   450
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   300
      Picture         =   "FormWB.frx":EBFB
      Top             =   6720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   9990
      Width           =   15225
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   13710
      Picture         =   "FormWB.frx":F054
      Top             =   6750
      Width           =   1155
   End
   Begin VB.Image ImgWB 
      Height          =   4095
      Left            =   1740
      Picture         =   "FormWB.frx":F4AD
      Top             =   6180
      Width           =   11445
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1785
      Index           =   10
      Left            =   11280
      Top             =   1830
      Width           =   3945
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   0
      Left            =   11280
      Top             =   1380
      Width           =   3945
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   5
      Left            =   5790
      Top             =   5010
      Width           =   3945
   End
End
Attribute VB_Name = "FormWB"
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

Private Sub Label3_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 20
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = DB460.Word(22)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 100# Then temp = 100#
        If temp < 20 Then temp = 20#
        DB460.Word(22) = temp
    End If
    Me.Update
End Sub

Private Sub TimerLocale_Timer()
      Me.Update
End Sub

Public Sub Update()
    Dim i As Integer
    Dim mask As Integer
    
    Label7 = Format(DB415.DWord(34) / 1000, "###0.000")
   ' VelocWBDisplay.Caption = DB460.Word(18)
    VelocLentaDisplay.Caption = DB460.Word(20)
    VelocRapidaDisplay.Caption = DB460.Word(22)
    Label3 = VelocRapidaDisplay
      ' aggiorna i dati pagina
    lblbar(2) = PaginaWb.Ordine_Descrizione
    lblbar(4) = PaginaWb.Ricetta_Descrizione
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    
'       For i = 0 To 15
'            ImgTubo(i).Visible = DB415.MaskBit(26, i)
'            ImgTubo(31 - i).Visible = DB415.MaskBit(32, i)
'            ImgTubo(47 - i).Visible = DB415.MaskBit(28, i)
'        Next i
          ' refresh colore tubi
          
    For i = 0 To 15
        mask = Abs(DB415.MaskBit(26, i)) Or 2 * Abs(DB415.MaskBit(28, i)) Or 4 * Abs(DB415.MaskBit(32, i))
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
        DisplayLunghezzaBarra.Caption = Unit.m_To_Display_mm(DB460.Word(4) / 1000#)
        'visualizza le dimensioni del tubo
        If DB460.Bit(2, 0) = True Then
     '      DisplayAltezza.Visible = False
     '      Label3(0).Visible = False
           DisplayDiametroBarra.Caption = Unit.m_To_Display_mm0(DB460.Word(6) / 10000#)
        Else
     '      DisplayAltezza.Visible = True
     '      Label3(0).Visible = True
           DisplayDiametroBarra.Caption = Unit.m_To_Display_mm0(DB460.Word(6) / 10000#)
        '   DisplayAltezza.Caption = Unit.m_To_Display_mm0(DB460.Word(8) / 10000#)
        End If
         
         If ControlloUpDownWB.Occupato = False Then
          ' If DB460.Word(34) < 20 Then DB460.Word(34) = 20
           ControlloUpDownWB.value = DB460.Word(18)
           ControlloUpDownWB.Refresh
        End If
         If ControlloMonobeam3.Occupato = False Then
          ' If DB460.Word(34) < 20 Then DB460.Word(34) = 20
           ControlloMonobeam3.value = DB460.Word(34)
           ControlloMonobeam3.Refresh
        End If
End Sub

'Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
'   frmKernel.PaginaCorrente = Index
'End Sub
''Private Sub DisplayAltezza_Click()
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
Private Sub DisplayDiametroBarra_Click()
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
'        DB460.Word(6) = temp
'        DB460.MaskBit(20, 0) = True     ' dati modificati, una modifica del diametro va propagata al resto della linea
'    End If
'    Me.Update
End Sub

Private Sub DisplayLunghezzaBarra_Click()
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
'        DB460.Word(4) = temp
'        DB460.MaskBit(20, 0) = True     ' dati modificati, una modifica della lunghezza va propagata al resto della linea
'    End If
'    Me.Update
End Sub
Private Sub Form_Activate()
    TimerLocale.Enabled = True
    TimerLocale.Interval = 500

    Barra21.Selezionato = 10
    WindowState = vbMaximized
      
      ' abilitazione temporizzatore locale
    TimerLocale.Enabled = True
    ControlloMonobeam3.Step = 10
    ControlloMonobeam3.LimMax = 100
    ControlloMonobeam3.LimMin = 20
    ControlloMonobeam3.Decimali = 0
    ControlloMonobeam3.Refresh
    ControlloUpDownWB.Step = 5
    ControlloUpDownWB.LimMax = 100
    ControlloUpDownWB.LimMin = 20
    ControlloUpDownWB.Decimali = 0
    ControlloUpDownWB.Refresh
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

Private Sub VelocLentaDisplay_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 20
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = DB460.Word(20)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 100# Then temp = 100#
        If temp < 20 Then temp = 20#
        DB460.Word(20) = temp
    End If
    Me.Update
End Sub
'
Private Sub VelocRapidaDisplay_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 20
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = DB460.Word(22)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 100# Then temp = 100#
        If temp < 20 Then temp = 20#
        DB460.Word(22) = temp
    End If
    Me.Update
End Sub
'
Private Sub VelocWBDisplay_Click()
    Dim temp As Double
    TOUCHNumericPad.Decimali = 0
    TOUCHNumericPad.ValoreMin = 20
    TOUCHNumericPad.ValoreMax = 100
    TOUCHNumericPad.Dati = DB460.Word(18)
    TOUCHNumericPad.Show vbModal
    If TOUCHNumericPad.DatiConfermati Then
        temp = TOUCHNumericPad.Dati
        If temp > 100# Then temp = 100#
        If temp < 20 Then temp = 20#
        DB460.Word(18) = temp
    End If
    Me.Update
End Sub


Private Sub Form_Load()
   
    ScritteMultilingua
  
   For i = 0 To 15
     ' picLogo(i).Visible = False
      Scritta(i).DrawingObject = picLogo(i)
      Scritta(i).StartColor = &HFFFF00   '&HC0C000   '&HFF00FF
      Scritta(i).EndColor = &HFFFFFF
   Next
        
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
 '  LblData.Top = 240
 '  LblOra.Top = 480
 '  LblOra.Left = 12480
 '  LblData.Left = 12000
   
   For i = 0 To 47
      If i < 16 Then ImgTubo(i).Picture = LoadPicture("..\bitmap\tubo.gif")
      If i > 15 And i < 32 Then ImgTubo(i).Picture = LoadPicture("..\bitmap\TuboRosso.gif")
      If i > 31 Then ImgTubo(i).Picture = LoadPicture("..\bitmap\tuboverde.gif")
   Next
   
   ImgTubo(48).Picture = LoadPicture("..\bitmap\tubo.gif")
   ImgTubo(49).Picture = LoadPicture("..\bitmap\tuboverde.gif")
   ImgTubo(50).Picture = LoadPicture("..\bitmap\tuborosso.gif")
   ImgWB.Picture = LoadPicture("..\bitmap\WB2.gif")
   WindowState = 2
End Sub
Public Sub AggiornaDaUpDownControl()
    DB460.Word(18) = ControlloUpDownWB.value
    DB460.Word(34) = ControlloMonobeam3.value
End Sub

Sub ScritteMultilingua()
    Label2.Caption = Param.Text("WB speed")
    LabelDiametroBarra.Caption = Param.Text("Diametro")
    LabelLungezzaBarra.Caption = Param.Text("Lunghezza")
 '   Label4.Caption = Param.Text("Low speed")
 '   Label6.Caption = Param.Text("Hi speed")
 '   Label7.Caption = Param.Text("Allineamento")
    Label9.Caption = Param.Text("VelMonobeam") & " 3 (%)"
     '======================================================================
    Scritta(15).RientroVert = 16
    Scritta(15).Caption = Param.Text("Prelievo")
    Scritta(14).RientroVert = 16
    Scritta(14).Caption = Param.Text("VR_Vel") & " 1"
    Scritta(13).RientroVert = 16
    Scritta(13).Caption = Param.Text("Filett.") & " 1"
    Scritta(12).RientroVert = 16
    Scritta(12).Caption = Param.Text("Sosta")
    Scritta(11).RientroVert = 16
    Scritta(11).Caption = Param.Text("Vernic.") & " 1"
    Scritta(10).RientroVert = 16
    Scritta(10).Caption = Param.Text("Sosta")
    Scritta(9).RientroVert = 16
    Scritta(9).Caption = Param.Text("Tappatrice")
    Scritta(8).RientroVert = 16
    Scritta(8).Caption = Param.Text("VR_Vel") & " 2"
    Scritta(7).RientroVert = 16
    Scritta(7).Caption = Param.Text("VR_Vel") & " 3"
    Scritta(6).RientroVert = 16
    Scritta(6).Caption = Param.Text("Filett.") & " 2"
    Scritta(5).RientroVert = 16
    Scritta(5).Caption = Param.Text("Sosta")
    Scritta(4).RientroVert = 16
    Scritta(4).Caption = Param.Text("Vernic.") & " 2"
    Scritta(3).RientroVert = 16
    Scritta(3).Caption = Param.Text("Sosta")
    Scritta(2).RientroVert = 16
    Scritta(2).Caption = Param.Text("Manicottatrice")
    Scritta(1).RientroVert = 16
    Scritta(1).Caption = Param.Text("Spintore")
    Scritta(0).RientroVert = 16
    Scritta(0).Caption = Param.Text("Deposito")
    '======================================================================
    Label8.Caption = Param.Text("Presenza")
    Label10.Caption = Param.Text("Completo")
    Label11.Caption = Param.Text("Non_abilitato")
    lblbar(5) = Param.Text("Walking beam page")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
   ' Label8.Caption = Param.Text("LungTubo")
    lblbar(1) = Param.Text("Pagina")
    Label4 = Param.Text("Velmodifica")
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

