VERSION 5.00
Begin VB.Form frmStatistica 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   8595
      Left            =   3180
      TabIndex        =   12
      Top             =   1560
      Width           =   9435
      Begin dp6.XPButton XPButton1 
         Height          =   885
         Index           =   5
         Left            =   600
         TabIndex        =   13
         Top             =   1620
         Width           =   4125
         _ExtentX        =   7382
         _ExtentY        =   1561
         TxtText         =   "            Storico allarmi"
         TxtTop          =   35
         TxtLeft         =   20
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\prjdll15.ico"
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
         Index           =   6
         Left            =   600
         TabIndex        =   14
         Top             =   2790
         Width           =   4125
         _ExtentX        =   7382
         _ExtentY        =   1561
         TxtText         =   "             Storico pacchi"
         TxtTop          =   35
         TxtLeft         =   20
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\_moricons65.ico"
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
         Height          =   855
         Index           =   10
         Left            =   6630
         TabIndex        =   15
         Top             =   2760
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1508
         TxtText         =   "Canc. storico pacchi"
         TxtTop          =   35
         TxtLeft         =   5
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\shell32102.ico"
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
         Index           =   11
         Left            =   6660
         TabIndex        =   16
         Top             =   1620
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1561
         TxtText         =   "Canc. storico allarmi"
         TxtTop          =   35
         TxtLeft         =   5
         BTYPE           =   3
         IMGTOP          =   5
         IMGLEFT         =   5
         ICONA           =   "..\bitmap\icone\shell32102.ico"
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Opzioni"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2820
         TabIndex        =   17
         Top             =   420
         Width           =   4875
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   3930
         X2              =   6660
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   3660
         X2              =   7440
         Y1              =   3210
         Y2              =   3210
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   915
         Index           =   4
         Left            =   3300
         Top             =   210
         Width           =   3915
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   18
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Image Image2 
         Height          =   1050
         Left            =   3180
         Picture         =   "frmstatistica.frx":0000
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
         TabIndex        =   7
         Top             =   150
         Width           =   8985
      End
   End
   Begin VB.Timer TimerLocale 
      Interval        =   500
      Left            =   3690
      Top             =   7770
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
End
Attribute VB_Name = "frmStatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Risp

Private Sub Barra21_PulsantePremuto(ByVal Index As IndicePuls)
   frmKernel.PaginaCorrente = Index
End Sub

Private Sub TimerLocale_Timer()
   Me.Update
End Sub
Sub Update()
       
       ' aggiorna i dati pagina
    lblbar(2) = PaginaEntrata.Ordine_Descrizione
    lblbar(4) = PaginaEntrata.Ricetta_Descrizione
    ' aggiorna lo stato del pulsante comunicazione
    If frmKernel.StatoCom Then
       If XPButton1(2).Icona <> "..\bitmap\semaforoRosso.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoRosso.gif"
    Else
        If XPButton1(2).Icona <> "..\bitmap\semaforoVerde.gif" Then XPButton1(2).Icona = "..\bitmap\semaforoVerde.gif"
    End If
    
'
End Sub
'
'Private Sub Barra1_PulsantePremuto(ByVal Index As IndicePuls)
'   frmKernel.PaginaCorrente = Index
'End Sub


Private Sub Form_Activate()
   
   TimerLocale.Enabled = True
   TimerLocale.Interval = 500
   Barra21.Selezionato = 2
    ' abilitazione temporizzatore locale
   Call Update
End Sub

Private Sub Form_Deactivate()
   TimerLocale.Enabled = False
End Sub

Private Sub Form_Load()
    ScritteMultilingua
End Sub

Sub ScritteMultilingua()
    lblbar(5) = Param.Text("Statis")
    lblbar(3) = Param.Text("Ricette")
    lblbar(0) = Param.Text("ORDER")
    lblbar(1) = Param.Text("Pagina")
    XPButton1(5).TxtText = Param.Text("StAll")
    XPButton1(6).TxtText = Param.Text("StPack")
    XPButton1(10).TxtText = Param.Text("CancStPack")
    XPButton1(11).TxtText = Param.Text("CancStAll")
    Label8 = Param.Text("Opzioni")
End Sub

Private Sub XPButton1_Click(Index As Integer)

     Select Case Index
     Case 0
'            frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
'            frmAnalisiAllarmi.Show
'            frmConnecting.Hide
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
                .NomeFile = "Statistica_pagina.htm"
               .Top = 1030
               .Left = 0
               .Width = 15350
               .Height = 9430
               .webPreview.Move 0, 0, frmHelp.Width, frmHelp.Height
               .Show
           End With
     Case 4
     Case 5
           frmConnecting.ShowConnecting "Refreshing alarms log grid. Please wait...", Me
           DoEvents
           frmAnalisiAllarmi.Show
           DoEvents
           frmConnecting.Hide
           Unload frmConnecting
     Case 6
           frmKernel.PaginaCorrente = PagStorico
     Case 11
          Risp = MsgBoxA(0, "Alarm log delete. Sure ?", App.Title, MB_ICONQUESTION, MB_YESNO)
          If Risp = vbNo Then Exit Sub
          On Error Resume Next
          Kill "..\target\AlarmsLOG.txt"
     Case 10
          Dim cnn As New ADODB.Connection
          
          Risp = MsgBoxA(0, "Bundle log delete. Sure ?", App.Title, MB_ICONQUESTION, MB_YESNO)
          If Risp <> vbYes Then Exit Sub
          On Error Resume Next
          cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\History.mdb;Persist Security Info=False"
          cnn.Execute "DELETE * FROM Bundles"
          cnn.Close
          Set cnn = Nothing
  End Select
  Exit Sub
  
ErrorePercorso:
           MsgBox "Percorso file guida errato", vbExclamation
End Sub
