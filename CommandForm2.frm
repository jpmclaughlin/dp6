VERSION 5.00
Begin VB.Form CommandForm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "422"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   360
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin dp6.Barra2 Barra21 
      Height          =   1155
      Left            =   0
      TabIndex        =   19
      Top             =   10410
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   2037
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
         Index           =   1
         Left            =   1650
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
      Begin VB.Label LblData 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   14
         Top             =   300
         Width           =   1185
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
         Index           =   3
         Left            =   7260
         TabIndex        =   13
         Top             =   660
         Width           =   3495
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
         Left            =   5760
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.Line Line2 
         X1              =   10860
         X2              =   10860
         Y1              =   180
         Y2              =   1020
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2003"
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
         Left            =   9600
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblOra 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10920
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   5760
         X2              =   10860
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "Anno"
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
         Index           =   4
         Left            =   8580
         TabIndex        =   9
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lblbar 
         BackColor       =   &H00EE3959&
         BackStyle       =   0  'Transparent
         Caption         =   "2309"
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
         Height          =   465
         Index           =   2
         Left            =   7260
         TabIndex        =   8
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblbar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Impianto"
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
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Image Image3 
         Height          =   1050
         Left            =   3150
         Picture         =   "CommandForm.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Index           =   2
         Left            =   3180
         TabIndex        =   1
         Top             =   150
         Width           =   8985
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Impianto      2333"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   555
         Left            =   6060
         TabIndex        =   6
         Top             =   150
         Width           =   3825
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Anno                                 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   375
         Left            =   6390
         TabIndex        =   5
         Top             =   690
         Width           =   3375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Data ora"
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
         Index           =   0
         Left            =   10590
         TabIndex        =   4
         Top             =   240
         Width           =   3435
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "DP6.1"
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
         Index           =   1
         Left            =   10740
         TabIndex        =   3
         Top             =   660
         Width           =   3435
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3210
         Picture         =   "CommandForm.frx":208E
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2205
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   14100
      Top             =   1290
   End
   Begin VB.Timer TimerLocale 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   14610
      Top             =   1320
   End
   Begin VB.PictureBox Selector 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   4
      Left            =   5040
      Picture         =   "CommandForm.frx":411C
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   11640
      Width           =   540
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   400
      Left            =   8280
      TabIndex        =   22
      Top             =   9000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   422
      Left            =   13260
      TabIndex        =   23
      Top             =   4050
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   410
      Left            =   4320
      TabIndex        =   24
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   420
      Left            =   8640
      TabIndex        =   25
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin dp6.Allarme Allarme1 
      Height          =   1335
      Index           =   424
      Left            =   10680
      TabIndex        =   26
      Top             =   2370
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   20
      Left            =   5850
      Top             =   5010
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   19
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   18
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   17
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   16
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   15
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   14
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   13
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   12
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   11
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   10
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   9
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   8
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   7
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   6
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   4
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   3
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   2
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   1
      Left            =   5190
      Top             =   4710
      Width           =   165
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   4
      X1              =   12780
      X2              =   12090
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   3
      X1              =   10350
      X2              =   9630
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   10350
      TabIndex        =   55
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Reggiatura"
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
      Index           =   6
      Left            =   10470
      TabIndex        =   54
      Top             =   2940
      Width           =   1545
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   735
      Index           =   6
      Left            =   10350
      Top             =   2640
      Width           =   1755
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   2  'Dash
      Height          =   1155
      Index           =   1
      Left            =   10230
      Top             =   2430
      Width           =   1995
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   7380
      TabIndex        =   53
      Top             =   6450
      Width           =   345
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   52
      Top             =   2670
      Width           =   345
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   486
      Left            =   12780
      TabIndex        =   28
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   450
      Left            =   7350
      TabIndex        =   20
      Top             =   7650
      Width           =   345
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   2940
      Picture         =   "CommandForm.frx":49F9
      Stretch         =   -1  'True
      Top             =   3300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   2940
      Picture         =   "CommandForm.frx":4E3B
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Stoccaggio"
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
      Index           =   5
      Left            =   13140
      TabIndex        =   51
      Top             =   2850
      Width           =   1545
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   5
      Left            =   12780
      Top             =   2640
      Width           =   2205
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   2
      X1              =   8520
      X2              =   8520
      Y1              =   3360
      Y2              =   4650
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Trasportatori laterali"
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
      Index           =   4
      Left            =   7770
      TabIndex        =   50
      Top             =   2760
      Width           =   1545
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2265
      Index           =   0
      Left            =   12480
      Top             =   3660
      Width           =   2625
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Packpipe"
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
      Height          =   345
      Index           =   3
      Left            =   7740
      TabIndex        =   49
      Top             =   4830
      Width           =   1545
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   1
      X1              =   8520
      X2              =   8520
      Y1              =   5340
      Y2              =   6450
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Soffiatura"
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
      Height          =   345
      Index           =   2
      Left            =   7740
      TabIndex        =   48
      Top             =   6660
      Width           =   1545
   End
   Begin VB.Label lblZoneName 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Tubificio"
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
      Height          =   675
      Index           =   1
      Left            =   240
      TabIndex        =   47
      Top             =   7860
      Width           =   1125
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   795
      Index           =   1
      Left            =   0
      Shape           =   2  'Oval
      Top             =   7620
      Width           =   1485
   End
   Begin VB.Label Label2 
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
      Left            =   300
      TabIndex        =   46
      Top             =   2340
      Width           =   3105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   465
      Index           =   0
      Left            =   240
      Top             =   2250
      Width           =   3285
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   5
      Left            =   9990
      Top             =   6570
      Width           =   165
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   12120
      TabIndex        =   45
      Top             =   8100
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   10170
      TabIndex        =   44
      Top             =   6630
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "SB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   540
      TabIndex        =   42
      Top             =   2850
      Width           =   405
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   13500
      TabIndex        =   41
      Top             =   7290
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   11850
      TabIndex        =   40
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   12480
      TabIndex        =   39
      Top             =   7740
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   12180
      TabIndex        =   38
      Top             =   7710
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   11850
      TabIndex        =   37
      Top             =   8070
      Width           =   255
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   13950
      TabIndex        =   36
      Top             =   7860
      Width           =   375
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P1"
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
      Left            =   12390
      TabIndex        =   35
      Top             =   7260
      Width           =   435
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P2"
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
      Left            =   12420
      TabIndex        =   34
      Top             =   6690
      Width           =   435
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P3"
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
      Left            =   12960
      TabIndex        =   33
      Top             =   7830
      Width           =   435
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "P4"
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
      Left            =   13440
      TabIndex        =   32
      Top             =   6810
      Width           =   435
   End
   Begin VB.Label EmNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "PE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   360
      TabIndex        =   31
      Top             =   3420
      Width           =   375
   End
   Begin VB.Label BarrNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   14460
      TabIndex        =   29
      Top             =   7050
      Width           =   255
   End
   Begin VB.Image Armadio 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   7920
      Stretch         =   -1  'True
      Top             =   9090
      Width           =   2175
   End
   Begin VB.Label LblCodOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   470
      Left            =   7380
      TabIndex        =   27
      Top             =   4650
      Width           =   345
   End
   Begin VB.Image imgDir 
      Height          =   480
      Index           =   6
      Left            =   1470
      Picture         =   "CommandForm.frx":527D
      Top             =   7800
      Width           =   480
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1470
      X2              =   7350
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      Index           =   0
      X1              =   8520
      X2              =   8520
      Y1              =   7140
      Y2              =   7650
   End
   Begin VB.Label lblZoneName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "via rulli"
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
      Height          =   345
      Index           =   0
      Left            =   7710
      TabIndex        =   21
      Top             =   7860
      Width           =   1545
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   0
      Left            =   7350
      Top             =   7650
      Width           =   2205
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   885
      Index           =   410
      Left            =   2220
      Top             =   7560
      Width           =   7875
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   1
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   5490
      Width           =   645
   End
   Begin VB.Shape BarrPallino 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      Height          =   525
      Index           =   0
      Left            =   300
      Top             =   2760
      Width           =   225
   End
   Begin VB.Shape EmPallino 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      Height          =   525
      Index           =   0
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3330
      Width           =   645
   End
   Begin VB.Label Commento 
      BackStyle       =   0  'Transparent
      Caption         =   "Safety door number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1050
      TabIndex        =   43
      Top             =   2910
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pulsanti emergenza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1050
      TabIndex        =   30
      Top             =   3420
      Width           =   3105
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   1185
      Index           =   10
      Left            =   240
      Top             =   2700
      Width           =   3285
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   2
      Left            =   7380
      Top             =   6450
      Width           =   2205
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1425
      Index           =   415
      Left            =   5070
      Top             =   6090
      Width           =   5025
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   3
      Left            =   7380
      Top             =   4650
      Width           =   2205
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2325
      Index           =   420
      Left            =   6690
      Top             =   3720
      Width           =   3405
   End
   Begin VB.Shape Flusso 
      BackStyle       =   1  'Opaque
      Height          =   705
      Index           =   4
      Left            =   7440
      Top             =   2670
      Width           =   2205
   End
   Begin VB.Shape LabelStato 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1365
      Index           =   422
      Left            =   6690
      Top             =   2310
      Width           =   8415
   End
End
Attribute VB_Name = "CommandForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PresentazioneOff As Boolean
Private Barriera(10) As Boolean
Private Emergenza(10) As Boolean

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub




Private Sub Timer1_Timer()
   Dim i As Integer
   
   DB400.WORDSReadAsync 18, 3
   TestBarriere
End Sub

'Private Sub Timer1_Timer()

'End Sub

'Private Sub Timer1_Timer()
'Dim i
'   List1.Clear
'   DB470.BlockReadAsync
'   For i = 1 To DB470.NumItems
'     List1.AddItem DB470.MultiReadValore(i)
'  Next
'   DB480.BlockReadAsync
'   For i = 1 To DB480.NumItems
'     List1.AddItem DB480.MultiReadValore(i)
'  Next
'  DB486.BlockReadAsync
'   For i = 1 To DB486.NumItems
'     List1.AddItem DB486.MultiReadValore(i)
'  Next
'End Sub

'=======================================================================
' CAMBIA IL VALORE DELLA VARIABILE CHE INDICA AL KENERL LA PRESENZA DI ALLARMI
'=======================================================================

Private Sub TimerLocale_Timer()
    Static oneShot As Integer
    Dim AllOn As Boolean
    Dim i As Integer
    Dim p As Boolean
    Dim ctl As Control
    Dim a As Integer
    
    On Error Resume Next
   
    For i = Allarme1.LBound To Allarme1.UBound
       If Allarme1(i).AllarmeTipo <> Nessuno Then Allarme1(i).RefreshTimer
    Next
    Me.Update
    On Error GoTo 0
    
   ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
  
End Sub
   
'==================================================================
' FUNZIONE DI AGGIORNAMENTO DELLO STATO DELLA PAGINA
'==================================================================

Public Sub Update()
    Dim i As Integer
    
'==================================================================
' VISUALIZZA GLI ALLARMI
'==================================================================
    
    'led di allarme generale

     AllarmiLayout DB410, 410
  '   AllarmiLayout DB411, 411
     AllarmiLayout DB400, 400
  '   AllarmiLayout DB412, 412
 '    AllarmiLayout DB413, 413
  '   AllarmiLayout DB414, 414
  '   AllarmiLayout DB415, 415
  '   AllarmiLayout DB416, 416
  '   AllarmiLayout DB417, 417
     AllarmiLayout DB420, 420
     AllarmiLayout DB424, 424
     AllarmiLayout DB422, 422
  '   AllarmiLayout DB425, 425

'     AllarmiLayout DB419, manicottatrice
'     AllarmiLayout DB418, tappatrice
''
'==================================================================
' VISUALIZZA LO STATO DELLE ZONE
'==================================================================
    ' stoccaggio
   ' ColoreLayout DB422, 0
    ' reggiatura
  '  ColoreLayout DB424, 1
    ' filettatura
  '   ColoreLayout DB416, 2
  '   ColoreLayout DB417, 3
    ' MPS
'=   ColoreLayout DB420, 4
    'ingresso WB
  '  ColoreLayout DB413, 3
    'entrata
     ColoreLayout DB410, 410
   ' ColoreLayout DB410, 3
   ' ColoreLayout DB410, 5
    ' trasportatori laterali
'=    ColoreLayout DB422, 8
'=    ColoreLayout DB422, 0
'=    ColoreLayout DB422, 2
'=    ColoreLayout DB422, 3
    ' uscita WB
'    ColoreLayout DB422, 2
    'via rulli 1
  '  ColoreLayout DB412, 10
     'bypass
 '   ColoreLayout DB412, 9
 '   ColoreLayout DB412, 7
    ' WB
  '  ColoreLayout DB415, 11
    
'==================================================================
' VISUALIZZA STATO BYPASS e filettatrici
'==================================================================
    
    
    'filettatrici

  '  If DB460.Bit(78, 0) = True Then
  '       Zona(7).Visible = True
  '       Zona(3).Visible = False
  '  Else
  '       Zona(7).Visible = False
  '       Zona(3).Visible = True
  '  End If
'
'    If DB460.Bit(118, 0) = True Then
'         Zona(4).Visible = False
'         Zona(8).Visible = True
'    Else
'         Zona(4).Visible = True
'         Zona(8).Visible = False
'    End If

'==================================================================
' VISUALIZZA L'ORDINE
'==================================================================
   
    LblCodOrd(470).Caption = frmKernel.CodOrdineCorrente.CodPacco
'    LblCodOrd(480).Caption = frmKernel.CodOrdineCorrente.CodRegge
    LblCodOrd(486).Caption = frmKernel.CodOrdineCorrente.CodStoccaggio
'    LblCodOrd(460).Caption = frmKernel.CodOrdineCorrente.CodWB
    LblCodOrd(450).Caption = frmKernel.CodOrdineCorrente.CodEntrata
    
    'bypass
'    If Param.GetBit("Par212_AttivaGestioneBypass") = False Then Exit Sub
'    If DB450.Bit(22, 0) = True And DB450.Bit(22, 1) = False And DB450.Bit(22, 2) = True And DB450.Bit(22, 3) = False Then
'       FrecciaDx.Visible = False
'       FrecciaGiu(0).Visible = True
'       FrecciaGiu(1).Visible = True
'       FrecciaGiu(2).Visible = False
'       FrecciaSx(0).Visible = False
'       FrecciaSx(1).Visible = False
'       Bypass(0).Visible = True
'       Bypass(1).Visible = False
'       ViaARulli(0).Visible = False
'       ViaARulli(1).Visible = False
'       ViaARulli(2).Visible = False
'       ViaARulli(3).Visible = False
'       LabelStato(14).Visible = True
'       LabelStato(15).Visible = True
'       LabelStato(16).Visible = False
'    End If
    
'  If DB450.Bit(22, 0) = True And DB450.Bit(22, 1) = False And DB450.Bit(22, 2) = False And DB450.Bit(22, 3) = True Then
'       FrecciaDx.Visible = True
'       FrecciaGiu(0).Visible = False
'       FrecciaGiu(1).Visible = True
'       FrecciaGiu(2).Visible = True
'       FrecciaSx(0).Visible = True
'       FrecciaSx(1).Visible = False
'       Bypass(0).Visible = True
'       Bypass(1).Visible = False
'       ViaARulli(0).Visible = True
'       ViaARulli(1).Visible = True
'       ViaARulli(2).Visible = True
'       ViaARulli(3).Visible = True
'       LabelStato(14).Visible = False
'       LabelStato(15).Visible = False
'       LabelStato(16).Visible = False
'    End If
    
'     If DB450.Bit(22, 0) = False And DB450.Bit(22, 1) = True And DB450.Bit(22, 2) = True And DB450.Bit(22, 3) = False Then
'       FrecciaDx.Visible = False
'       FrecciaGiu(0).Visible = True
'       FrecciaGiu(1).Visible = True
'       FrecciaGiu(2).Visible = False
'       FrecciaSx(0).Visible = False
'       FrecciaSx(1).Visible = True
'       Bypass(0).Visible = False
'       Bypass(1).Visible = True
'       ViaARulli(0).Visible = True
'       ViaARulli(1).Visible = True
'       ViaARulli(2).Visible = False
 '      ViaARulli(3).Visible = False
'       LabelStato(14).Visible = False
'       LabelStato(15).Visible = True
'       LabelStato(16).Visible = False
'    End If
    
'     If DB450.Bit(22, 0) = False And DB450.Bit(22, 1) = True And DB450.Bit(22, 2) = False And DB450.Bit(22, 3) = True Then
'       FrecciaDx.Visible = False
'       FrecciaGiu(0).Visible = False
'       FrecciaGiu(1).Visible = True
'       FrecciaGiu(2).Visible = True
'       FrecciaSx(0).Visible = True
'       FrecciaSx(1).Visible = False
'       Bypass(0).Visible = False
'       Bypass(1).Visible = True
'       ViaARulli(0).Visible = False
'       ViaARulli(1).Visible = False
'       ViaARulli(2).Visible = True'
'       ViaARulli(3).Visible = True
'       LabelStato(14).Visible = False
'       LabelStato(15).Visible = False
'       LabelStato(16).Visible = True
'    End If
    
End Sub


''==================================================================
'' FUNZIONE ABILITAZIONE ALLARMI
''==================================================================

Private Sub AllarmiLayout(DBSource As DBClass, Index As Integer)
    If DBSource.MaskBit(0, 8) = True And DBSource.MaskBit(0, 9) = True Then
       Allarme1(Index).AllarmeTipo = Entrambi
    Else
        If DBSource.MaskBit(0, 8) = True And DBSource.MaskBit(0, 9) = False Then
            Allarme1(Index).AllarmeTipo = Allarme
        Else
            If DBSource.MaskBit(0, 8) = False And DBSource.MaskBit(0, 9) = True Then
                 Allarme1(Index).AllarmeTipo = Messaggio
            Else
                If DBSource.MaskBit(0, 8) = False And DBSource.MaskBit(0, 9) = False Then
                    Allarme1(Index).AllarmeTipo = Nessuno
                End If
            End If
        End If
    End If
End Sub

'==================================================================
' FUNZIONE COLORAZIONE LAYOUT
'==================================================================

Private Sub ColoreLayout(DBSource As DBClass, Index As Integer)
        If DBSource.MaskBit(0, 2) Then
            LabelStato(Index).BackColor = ManualColor
        Else
            If DBSource.MaskBit(0, 3) Then
                LabelStato(Index).BackColor = SemiautoColor
            Else
                If DBSource.MaskBit(0, 4) Then
                    LabelStato(Index).BackColor = AutoColor
                Else
                    LabelStato(Index).BackColor = E_StopColor
                End If
            End If
        End If
End Sub

'==================================================================
' EVENTO FOCUS AL FORM
'==================================================================

Public Sub Form_Activate()
 ' Barra1.Pulsante_Click 3
  Barra21.Selezionato = 3
  If PresentazioneOff = False Then
    TimerLocale.Enabled = False
    Timer1.Enabled = False
  '  Barra1.Bloccata = True
  Else
    Call Aggiornamento
  End If
End Sub
Sub Aggiornamento()
    frmKernel.PulAllarmiPremuto = False
    TimerLocale.Enabled = True
    Timer1.Enabled = True
    Timer1.Interval = 500
    TimerLocale.Interval = 500
    Me.Update
 '   Barra1.Bloccata = False
    TestBarriere
End Sub
Private Sub Form_Deactivate()
    TimerLocale.Enabled = False
    Timer1.Enabled = False
End Sub

'==================================================================
' EVENTO CARICA RISORSE NEL FORM
'==================================================================

Private Sub Form_Load()
Dim i As Integer

'   posizione data e ora
  ' LblData.Top = 240
'   LblOra.Top = 480
'   LblOra.Left = 12480
 '  LblData.Left = 12000
   
  '  Picture = LoadPicture("..\bitmap\SfondoDP6_ver2_0.jpg")
    WindowState = 2
    
    'frequenza lampeggio allarmi
    
    On Error Resume Next
    For i = Allarme1.LBound To Allarme1.UBound
       Allarme1(i).AllarmeTipo = Nessuno
       Allarme1(i).Intervallo = 500
    Next
    On Error GoTo 0
    
    'carica le immagini per le varie zone
  '  Picture1 = LoadPicture("..\bitmap\layout.gif")
  '  Zona(420) = LoadPicture("..\bitmap\packpipe.gif")
 '   Zona(422) = LoadPicture("..\bitmap\stoccaggio.gif")
'    ImgFreccia = LoadPicture("..\bitmap\FrecciaSxByPass.gif")
'    Zona(2) = LoadPicture("..\bitmap\WB.gif")
'    Zona(3) = LoadPicture("..\bitmap\filettatriceSxOff.gif")
'    Zona(4) = LoadPicture("..\bitmap\filettatriceDxOff.gif")
'    Zona(5) = LoadPicture("..\bitmap\reggsignode.gif")
  '  Zona(2) = LoadPicture("..\bitmap\Entrata.gif")
'    Zona(7) = LoadPicture("..\bitmap\filettatriceSx.gif")
'    Zona(8) = LoadPicture("..\bitmap\filettatriceDx.gif")
'   Zona(410) = LoadPicture("..\bitmap\entrata.gif")
'   Pulpito = LoadPicture("..\bitmap\Pulpito90.gif")
   Armadio = LoadPicture("..\bitmap\ArmadioElettrico.gif")
 ' StatoPLC(0) = LoadPicture("..\bitmap\semaforoRosso.gif")
 '  StatoPLC(1) = LoadPicture("..\bitmap\semaforoVerde.gif")
   'Bypass(0) = LoadPicture("..\bitmap\puntodibypass.gif")
   'Bypass(1) = LoadPicture("..\bitmap\puntodibypass.gif")
   'FrecciaGiu(0) = LoadPicture("..\bitmap\FrecciaGiuByPass.gif")
   'FrecciaGiu(1) = LoadPicture("..\bitmap\FrecciaGiuByPass.gif")
   'FrecciaGiu(2) = LoadPicture("..\bitmap\FrecciaGiuByPass.gif")
   'FrecciaDx = LoadPicture("..\bitmap\FrecciaDxByPass.gif")
   'FrecciaSx(0) = LoadPicture("..\bitmap\FrecciaSxByPass.gif")
   'FrecciaSx(1) = LoadPicture("..\bitmap\FrecciaSxByPass.gif")
'   ViaARulli(0) = LoadPicture("..\bitmap\RullieraVistaAlto.gif")
'   ViaARulli(1) = LoadPicture("..\bitmap\RullieraVistaAlto.gif")
   'ViaARulli(2) = LoadPicture("..\bitmap\RullieraVistaAlto.gif")
   'ViaARulli(3) = LoadPicture("..\bitmap\RullieraVistaAlto.gif")
   'Accumolo = LoadPicture("..\bitmap\accumolov.gif")
   
   ScritteMultilingua
  
   
End Sub

'==================================================================
' VISUALIZZA LA PAGINA DEGLI ALLARMI IN BASE ALL'ALLARME CLICCATO
'==================================================================

Private Sub Allarme1_Cliccato(Index As Integer)
     
    frmKernel.PulAllarmiPremuto = True
    'inizializzazione allarmi

    AlarmForm.CheckDB400.Value = 0
    AlarmForm.CheckDB410.Value = 0
'    AlarmForm.CheckDB411.value = 0
'    AlarmForm.CheckDB412.value = 0
'    AlarmForm.CheckDB413.Value = 0
'    AlarmForm.CheckDB414.value = 0
'    AlarmForm.CheckDB415.value = 0
'    AlarmForm.CheckDB416.value = 0
'    AlarmForm.CheckDB417.value = 0
'    AlarmForm.CheckDB418.value = 0
'    AlarmForm.CheckDB419.value = 0
    AlarmForm.CheckDB420.Value = 0
    AlarmForm.CheckDB422.Value = 0
    AlarmForm.CheckDB424.Value = 0
  '  AlarmForm.CheckDB425.value = 0
    
   'Setta l'allarme selezionato
   
    Select Case Index
        Case 400
            AlarmForm.CheckDB400.Value = 1
        Case 410
            AlarmForm.CheckDB410.Value = 1
        Case 411
            AlarmForm.CheckDB411.Value = 1
        Case 412
            AlarmForm.CheckDB412.Value = 1
        Case 413
            AlarmForm.CheckDB413.Value = 1
        Case 414
            AlarmForm.CheckDB414.Value = 1
        Case 415
            AlarmForm.CheckDB415.Value = 1
        Case 416
            AlarmForm.CheckDB416.Value = 1
        Case 417
            AlarmForm.CheckDB417.Value = 1
        Case 418
            AlarmForm.CheckDB418.Value = 1
        Case 419
            AlarmForm.CheckDB419.Value = 1
        Case 420
            AlarmForm.CheckDB420.Value = 1
        Case 422
            AlarmForm.CheckDB422.Value = 1
        Case 424
            AlarmForm.CheckDB424.Value = 1
        Case 425
            AlarmForm.CheckDB425.Value = 1
    End Select
    
    'visualizza il form allarmi
    
    AlarmForm.ZOrder 0
    AlarmForm.Visible = True
    AlarmForm.WindowState = vbMaximized
    
End Sub

Private Sub Label5_Click()
   ErrDBPiccolo = False
End Sub

Sub ScritteMultilingua()
   lblbar(0) = Param.Text("Impianto")
   lblbar(2) = Param.Text("Comm_Num")
   lblbar(4) = Param.Text("Anno")
   lblbar(5) = Param.Text("Comm_Anno")
   lblbar(1) = Param.Text("Pagina")
   Commento.Caption = Param.Text("LegBarr")
   Label1 = Param.Text("PulEm")
 '  lblZoneName(410) = Param.Text("ViaRulli")
 '  lblZoneName(420) = Param.Text("PP")
 '  lblZoneName(422) = Param.Text("Trlat")
 '  lblZoneName(425) = Param.Text("Pesatura")
 '  lblZoneName(424) = Param.Text("Reggiatura")
 '  lblZoneName(426) = Param.Text("Stoccaggio")
End Sub

Sub TestBarriere()
    Dim i As Integer
    Dim Br As Boolean
    Dim Em As Boolean
    Static Lamp As Boolean
    
    Lamp = Not Lamp
    On Error Resume Next
    DoEvents
    
    Barriera(1) = DB400.MaskBit(20, 0)
    Barriera(2) = DB400.MaskBit(20, 1)
    Barriera(3) = DB400.MaskBit(20, 2)
    Barriera(4) = DB400.MaskBit(20, 3)
    Barriera(5) = DB400.MaskBit(20, 4)
    Barriera(6) = DB400.MaskBit(20, 5)
    Barriera(7) = DB400.MaskBit(20, 6)
    'Barriera(8) = DB400.MaskBit(20, 7)
    Barriera(9) = DB400.MaskBit(20, 8)
    Barriera(10) = DB400.MaskBit(20, 9)
    
    Emergenza(0) = DB400.MaskBit(18, 0)
    Emergenza(1) = DB400.MaskBit(18, 1)
    Emergenza(2) = DB400.MaskBit(18, 2)
    Emergenza(3) = DB400.MaskBit(18, 3)
'    Emergenza(4) = DB400.MaskBit(18, 0)
'    Emergenza(5) = DB400.MaskBit(18, 0)
'    Emergenza(6) = DB400.MaskBit(18, 0)
'    Emergenza(7) = DB400.MaskBit(18, 0)
'    Emergenza(8) = DB400.MaskBit(18, 0)
'    Emergenza(9) = DB400.MaskBit(18, 0)
'    Emergenza(10) = DB400.MaskBit(18, 0)
    
    Br = False: Em = False
    For i = 1 To 10
       BarrPallino(i).Visible = Barriera(i) And Lamp
       BarrNumero(i).Visible = BarrPallino(i).Visible
       EmPallino(i).Visible = Emergenza(i) And Lamp
       EmNumero(i).Visible = EmPallino(i).Visible
       Br = Br Or BarrPallino(i).Visible
       Em = Em Or EmPallino(i).Visible
    Next
    Image1.Visible = Br
    Image4.Visible = Em

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
           frmHelp.Errori = False
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
                .NomeFile = "mappa_pagina.htm"
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

