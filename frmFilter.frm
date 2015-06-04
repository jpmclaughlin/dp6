VERSION 5.00
Begin VB.Form frmFilter 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recipes filter"
   ClientHeight    =   8715
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   6060
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tube"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3735
      Left            =   180
      TabIndex        =   20
      Top             =   570
      Width           =   5715
      Begin VB.TextBox minTubew 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   31
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3870
         TabIndex        =   30
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   29
         Top             =   2010
         Width           =   1365
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3870
         TabIndex        =   28
         Top             =   2040
         Width           =   1365
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   27
         Top             =   2610
         Width           =   1365
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3870
         TabIndex        =   26
         Top             =   2610
         Width           =   1365
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   25
         Top             =   3180
         Width           =   1365
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3870
         TabIndex        =   24
         Top             =   3180
         Width           =   1365
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Round"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   450
         TabIndex        =   23
         Top             =   390
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Square"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   2310
         TabIndex        =   22
         Top             =   390
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   4170
         TabIndex        =   21
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4260
         TabIndex        =   37
         Top             =   1050
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   510
         TabIndex        =   36
         Top             =   1050
         Width           =   525
      End
      Begin VB.Label tubew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tube width"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1860
         TabIndex        =   35
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label tubeh 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tube height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1830
         TabIndex        =   34
         Top             =   2070
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tube lenght"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1860
         TabIndex        =   33
         Top             =   2670
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tube thickn."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1860
         TabIndex        =   32
         Top             =   3270
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bundle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3495
      Left            =   180
      TabIndex        =   4
      Top             =   4380
      Width           =   5715
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   15
         Top             =   1230
         Width           =   1365
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         TabIndex        =   14
         Top             =   1230
         Width           =   1365
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1365
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         TabIndex        =   12
         Top             =   1800
         Width           =   1365
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   11
         Top             =   2370
         Width           =   1365
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         TabIndex        =   10
         Top             =   2370
         Width           =   1365
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   2910
         Width           =   1365
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3900
         TabIndex        =   8
         Top             =   2910
         Width           =   1365
      End
      Begin VB.OptionButton Pacco 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Esa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Index           =   0
         Left            =   420
         TabIndex        =   7
         Top             =   330
         Width           =   1515
      End
      Begin VB.OptionButton Pacco 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Square"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Index           =   1
         Left            =   2280
         TabIndex        =   6
         Top             =   330
         Width           =   1515
      End
      Begin VB.OptionButton Pacco 
         BackColor       =   &H00C0C0C0&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   555
         Index           =   2
         Left            =   4140
         TabIndex        =   5
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4230
         TabIndex        =   39
         Top             =   870
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   480
         TabIndex        =   38
         Top             =   870
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tubes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2130
         TabIndex        =   19
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Row base tubes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1590
         TabIndex        =   18
         Top             =   1860
         Width           =   2250
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2190
         TabIndex        =   17
         Top             =   2430
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2190
         TabIndex        =   16
         Top             =   2970
         Width           =   990
      End
   End
   Begin VB.TextBox Text16 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2100
      TabIndex        =   3
      Top             =   60
      Width           =   3885
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7950
      Width           =   1770
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FF00&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7980
      Width           =   1770
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recipe name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   30
      TabIndex        =   2
      Top             =   60
      Width           =   1875
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tfiltro
           ttubo As Variant
           twmin As Variant
           twmax As Variant
           thmin As Variant
           thmax As Variant
           tlmin As Variant
           tlmax As Variant
           tsmin As Variant
           tsmax As Variant
           tpacco As Variant
           ptubesmin As Variant
           ptubesmax As Variant
           pfilemin As Variant
           pfilemax As Variant
           pbasemin As Variant
           pbasemax As Variant
           ppesomin As Variant
           ppesomax As Variant
End Type

Public Filtrostr As String
Private FiltroNome As String
Private App As Variant
Private Datifiltro As tfiltro

Private Sub cmdCancel_Click()
    Filtrostr = ""
    FiltroNome = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Filtrostr = ConcateFilter
    Unload Me
End Sub

Private Sub Form_Load()
   Option3.value = True
   Pacco(2).value = True
   Datifiltro.pbasemax = ""
   Datifiltro.pbasemin = ""
   Datifiltro.pfilemax = ""
   Datifiltro.pfilemin = ""
   Datifiltro.ppesomax = ""
   Datifiltro.ppesomin = ""
   Datifiltro.ptubesmax = ""
   Datifiltro.ptubesmin = ""
   Datifiltro.thmax = ""
   Datifiltro.thmin = ""
   Datifiltro.tlmax = ""
   Datifiltro.tlmin = ""
   Datifiltro.tpacco = ""
   Datifiltro.tsmax = ""
   Datifiltro.tsmin = ""
   Datifiltro.ttubo = ""
   Datifiltro.twmax = ""
   Datifiltro.twmin = ""
   FiltroNome = ""
End Sub

Private Sub Option1_Click()
    Text2.Visible = False
    Text3.Visible = False
    tubeh.Visible = False
    Datifiltro.ttubo = "1"
End Sub

Private Sub Option2_Click()
    Text2.Visible = True
    Text3.Visible = True
    tubeh.Visible = True
    Datifiltro.ttubo = "2"
End Sub

Private Sub Option3_Click()
    Text2.Visible = True
    Text3.Visible = True
    tubeh.Visible = True
    Datifiltro.ttubo = ""
End Sub
Private Sub minTubew_Click()
    On Error GoTo errore
    App = Immissione(minTubew)
    Datifiltro.twmin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text1_Click()
    On Error GoTo errore
    App = Immissione(Text1)
    Datifiltro.twmax = App
    Exit Sub
    
errore:
End Sub



Private Sub Text16_Click()
    TOUCHKeyBoard.Dati = FiltroNome
    TOUCHKeyBoard.Show vbModal
    If TOUCHKeyBoard.DatiConfermati Then
        FiltroNome = TOUCHKeyBoard.Dati
        Text16 = TOUCHKeyBoard.Dati
    End If
End Sub

Private Sub Text2_Click()
    On Error GoTo errore
    App = Immissione(Text2)
    Datifiltro.thmin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text3_Click()
    On Error GoTo errore
    App = Immissione(Text3)
    Datifiltro.thmax = App
    Exit Sub
    
errore:
End Sub

Private Sub Text4_Click()
    On Error GoTo errore
    App = Immissione(Text4)
    Datifiltro.tlmin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text5_Click()
    On Error GoTo errore
    App = Immissione(Text5)
    Datifiltro.tlmax = App
    Exit Sub
    
errore:
End Sub

Private Sub Text8_Click()
    On Error GoTo errore
    App = Immissione(Text8)
    Datifiltro.tsmin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text9_Click()
    On Error GoTo errore
    App = Immissione(Text9)
    Datifiltro.tsmax = App
    Exit Sub
    
errore:
End Sub
Private Sub Text6_Click()
    On Error GoTo errore
    App = Immissione(Text6)
    Datifiltro.ptubesmin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text7_Click()
    On Error GoTo errore
    App = Immissione(Text7)
    Datifiltro.ptubesmax = App
    Exit Sub
    
errore:
End Sub

Private Sub Text10_Click()
    On Error GoTo errore
    App = Immissione(Text10)
    Datifiltro.pbasemin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text11_Click()
    On Error GoTo errore
    App = Immissione(Text11)
    Datifiltro.pbasemax = App
    Exit Sub
    
errore:
End Sub
Private Sub Text12_Click()
    On Error GoTo errore
    App = Immissione(Text12)
    Datifiltro.pfilemin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text13_Click()
    On Error GoTo errore
    App = Immissione(Text13)
    Datifiltro.pfilemax = App
    Exit Sub
    
errore:
End Sub
Private Sub Text14_Click()
    On Error GoTo errore
    App = Immissione(Text14)
    Datifiltro.ppesomin = App
    Exit Sub
    
errore:
End Sub

Private Sub Text15_Click()
    On Error GoTo errore
    App = Immissione(Text15)
    Datifiltro.ppesomax = App
    Exit Sub
    
errore:
End Sub

Private Sub Pacco_Click(Index As Integer)
    If Index = 2 Then Datifiltro.tpacco = "": Exit Sub
    Datifiltro.tpacco = CStr(Index + 1)
End Sub

Function Immissione(inValore As TextBox) As Variant
            TOUCHNumericPad.ValoreMin = 0
            TOUCHNumericPad.ValoreMax = 100000
            TOUCHNumericPad.Dati = Val(inValore.Text)
            DoEvents
            TOUCHNumericPad.Show vbModal
            DoEvents
            If TOUCHNumericPad.DatiConfermati Then
                Immissione = TOUCHNumericPad.Dati
                inValore.Text = Immissione
                If Immissione = 0 Then inValore.Text = "": Immissione = ""
            End If
End Function

Function ConcateFilter() As String
   ConcateFilter = " WHERE" & IIf(Trim(Datifiltro.ttubo) <> "", " TipoTubo=" & Trim(Datifiltro.ttubo), " (TipoTubo=1 OR TipoTubo=2 OR TipoTubo=0)")
    If Datifiltro.twmin <> "" Then ConcateFilter = ConcateFilter & " AND Larghezza>=" & CSng(Datifiltro.twmin / 1000)
    If Datifiltro.twmax <> "" Then ConcateFilter = ConcateFilter & " AND Larghezza<=" & CSng(Datifiltro.twmax / 1000)
    If Datifiltro.thmin <> "" Then ConcateFilter = ConcateFilter & " AND Altezza>=" & CSng(Datifiltro.thmin / 1000)
    If Datifiltro.thmax <> "" Then ConcateFilter = ConcateFilter & " AND Altezza<=" & CSng(Datifiltro.thmax / 1000)
    If Datifiltro.tlmin <> "" Then ConcateFilter = ConcateFilter & " AND Lunghezza>=" & CSng(Datifiltro.tlmin / 1000)
    If Datifiltro.tlmax <> "" Then ConcateFilter = ConcateFilter & " AND Lunghezza<=" & CSng(Datifiltro.tlmax / 1000)
    If Datifiltro.tsmin <> "" Then ConcateFilter = ConcateFilter & " AND Spessore>=" & CSng(Datifiltro.tsmin / 1000)
    If Datifiltro.tsmax <> "" Then ConcateFilter = ConcateFilter & " AND Spessore<=" & CSng(Datifiltro.tsmax / 1000)
    If Datifiltro.ptubesmin <> "" Then ConcateFilter = ConcateFilter & " AND NumeroTubi>=" & CLng(Datifiltro.ptubesmin)
    If Datifiltro.ptubesmax <> "" Then ConcateFilter = ConcateFilter & " AND NumeroTubi<=" & CLng(Datifiltro.ptubesmax)
    If Datifiltro.pbasemin <> "" Then ConcateFilter = ConcateFilter & " AND Fila01>=" & CLng(Datifiltro.pbasemin)
    If Datifiltro.pbasemax <> "" Then ConcateFilter = ConcateFilter & " AND Fila01<=" & CLng(Datifiltro.pbasemax)
    If Datifiltro.pfilemin <> "" Then ConcateFilter = ConcateFilter & " AND NumeroFile>=" & CLng(Datifiltro.pfilemin)
    If Datifiltro.pfilemax <> "" Then ConcateFilter = ConcateFilter & " AND NumeroFile<=" & CLng(Datifiltro.pfilemax)
    If Datifiltro.ppesomin <> "" Then ConcateFilter = ConcateFilter & " AND PesoTeoricoPacco>=" & CSng(Datifiltro.ppesomin)
    If Datifiltro.ppesomax <> "" Then ConcateFilter = ConcateFilter & " AND PesoTeoricoPacco<=" & CSng(Datifiltro.ppesomax)
    If Datifiltro.tpacco <> "" Then ConcateFilter = ConcateFilter & " AND TipoPacco=" & Trim(Datifiltro.tpacco)
    If FiltroNome <> "" Then ConcateFilter = ConcateFilter & " AND ID LIKE '%" & FiltroNome & "%'"
    ConcateFilter = Replace(ConcateFilter, ",", ".")
End Function
