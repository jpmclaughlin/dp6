VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStampa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label preview"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7710
      Visible         =   0   'False
      Width           =   2115
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   7710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStampa.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStampa.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStampa.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStampa.frx":0A86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStampa.frx":0DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FormStampa.frx":10BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1770
      TabIndex        =   42
      Top             =   10860
      Width           =   2055
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   510
      TabIndex        =   41
      Top             =   10020
      Width           =   4665
   End
   Begin VB.Label Label40 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "WEIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2790
      TabIndex        =   40
      Top             =   9750
      Width           =   1125
   End
   Begin VB.Label Label39 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4890
      TabIndex        =   39
      Top             =   9750
      Width           =   495
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3810
      TabIndex        =   38
      Top             =   9750
      Width           =   1035
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "QTY PIECES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   37
      Top             =   9750
      Width           =   1695
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1710
      TabIndex        =   36
      Top             =   9750
      Width           =   1065
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LENGTH "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   35
      Top             =   9450
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1110
      TabIndex        =   34
      Top             =   9450
      Width           =   915
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2640
      TabIndex        =   33
      Top             =   9450
      Width           =   645
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2160
      TabIndex        =   32
      Top             =   9450
      Width           =   375
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3330
      TabIndex        =   31
      Top             =   9450
      Width           =   855
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1110
      TabIndex        =   30
      Top             =   9150
      Width           =   4425
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GRADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   29
      Top             =   9150
      Width           =   1215
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   840
      TabIndex        =   28
      Top             =   8850
      Width           =   4695
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SIZE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   27
      Top             =   8850
      Width           =   615
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "HEAT #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2550
      TabIndex        =   26
      Top             =   8550
      Width           =   1095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3690
      TabIndex        =   25
      Top             =   8550
      Width           =   1845
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   24
      Top             =   8550
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   840
      TabIndex        =   23
      Top             =   8550
      Width           =   1665
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   4350
      Width           =   855
   End
   Begin VB.Label Label36 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "FT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2310
      TabIndex        =   21
      Top             =   4350
      Width           =   375
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GRADE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3900
      Width           =   975
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1830
      TabIndex        =   19
      Top             =   4830
      Width           =   1635
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "SIZE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3420
      Width           =   615
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   7170
      Width           =   2055
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   510
      TabIndex        =   16
      Top             =   6300
      Width           =   4665
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2790
      TabIndex        =   14
      Top             =   4350
      Width           =   645
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1260
      TabIndex        =   13
      Top             =   4350
      Width           =   915
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1230
      TabIndex        =   12
      Top             =   3900
      Width           =   4245
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   3420
      Width           =   4515
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   2940
      Width           =   4155
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LBS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3570
      TabIndex        =   9
      Top             =   5280
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1320
      Index           =   1
      Left            =   1800
      Picture         =   "FormStampa.frx":13D4
      Stretch         =   -1  'True
      Top             =   780
      Width           =   2145
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "HEAT #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "QTY PIECES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4830
      Width           =   1695
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "LENGTH "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4350
      Width           =   1095
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "BUNDLE WEIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2460
      Width           =   4515
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5670
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   5670
      Width           =   3795
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1680
      X2              =   5460
      Y1              =   6030
      Y2              =   6030
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2460
      Width           =   735
   End
End
Attribute VB_Name = "frmStampa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx, yy

Private Type ctlDATI
        Nome As String * 100
        PosX As Long
        PosY As Long
        FontSize As Long
        FontName As String * 20
        FontBold As Boolean
        FontItal As Boolean
        Visibile As Boolean
End Type

Private Type DATItmp
        Nome As String * 100
        PosX As String * 10
        PosY As String * 10
        FontSize As String * 4
        FontName As String * 20
        FontBold As String * 2
        FontItal As String * 2
        Visibile As String * 2
End Type

Private ColAUTO As New Collection
Private Writetemp As DATItmp
Private Posizione As POINTAPI
Private CTLS(20) As ctlDATI
Private ManCount As Long
Private CtlSelected As Integer
Public cmdHoldShow As Boolean
Public cmdHoldFont As Boolean

Public Sub PrintExec(Optional ByVal ordercode As Integer, Optional ingrade = "")
     Dim OrderTemp As New OrderClass
     Dim RecipeTemp As New RecipeClass
     
     Dim i As Integer
     
     OrderTemp.UploadData ordercode
     RecipeTemp.UploadData OrderTemp.IDRicetta
     
     Command1.Visible = False
     
     Label1 = Date
     Label2 = Date
     Label18 = OrderTemp.CampoManuale(3) '& "," & OrderTemp.CampoManuale(4)
     If Trim(OrderTemp.CampoManuale(10)) <> "" Then
        Label18 = Label18 & "," & OrderTemp.CampoManuale(10)
     End If
     Label7 = Label18
     
     If Trim(OrderTemp.IDRicetta) = "" Then OrderTemp.IDRicetta = "0 X 0"
     Label21 = Left(UCase(OrderTemp.IDRicetta), InStrRev(UCase(OrderTemp.IDRicetta), "X") - 1) 'OrderTemp.IDRicetta
     Label11 = Label21
     Label22 = IIf(ingrade = "", RecipeTemp.Grade, ingrade)
     Label13 = Label22
     
'     If Round(Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.ft, 6), 1) - Fix(Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.ft, 6)) < 0.9 Then
'        Label24 = Fix(Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.ft, 6))
        Label24 = Right(OrderTemp.IDRicetta, Len(Trim(UCase(OrderTemp.IDRicetta))) - InStrRev(Trim(UCase(OrderTemp.IDRicetta)), "X")) 'Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.ft, 0)
        Label20 = Label24
'        Label28 = Round(Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.inch, 6) - Fix(Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.inch, 6)), 0)
'        Label17 = Label28
'        Label37 = "IN"
'        Label15 = Label37
'     Else
'        Label24 = Round(Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.ft, 6), 0)
'        Label24 = Conv_UM.Conversione(RecipeTemp.TuboLunghezza, UM.mt, UM.ft, 3)
'        Label20 = Label24
'        Label28 = ""
'        Label17 = Label28
'        Label37 = ""
'        Label15 = Label37
'     End If
     
     Label34 = RecipeTemp.NumeroTubiPacco
     Label27 = Label34
     Label30 = WeightToPrint '= Cartellino.PesoPaccoRecipeTemp.Weight
     Label38 = Label30
     Label4 = OrderTemp.CampoManuale(1)
     Label31 = ("*" & DB402.DWord(40) & ("*"))
     Label41 = Label31
     Label32 = DB402.DWord(40)
     Label42 = Label32
     
     For i = 1 To Param.GetNumber("Par221_NumeroCartellini")
        PrintForm
        Printer.EndDoc
        DoEvents
     Next
     
     Command1.Visible = True
     
End Sub

Sub RefreshVar()
   Dim i As Integer
   Dim manindex As Integer
   
   manindex = 0
   For i = 0 To 20
    ' If Cartellino.manIndice(i) <> 0 Then
     If DatiCartellino(i) <> 0 Then
        If i < 11 Then
          'Man(Cartellino.manIndice(i)) = Cartellino.CampoManuale(i)
'          Man(DatiCartellino(i)) = Cartellino.CampoManuale(i)
        Else
         ' Man(Cartellino.manIndice(i)) = Cartellino.CampoAuto(i)
'          Man(DatiCartellino(i)) = Cartellino.CampoAuto(i - 10)
        End If
     End If
   Next
End Sub
Function LoadFixTexts() As Boolean
  On Error Resume Next
  LoadFixTexts = False
  RS_LabelText.Close
  RS_LabelText.Open "..\Target\Tickets\Labeltexts.xml", , , , adCmdFile
  If RS_LabelText.EOF = False Then
     LoadFixTexts = True
  End If
  Exit Function
Errore:
End Function
Function LengFixTextRefresh(ByVal inLingua As TCampoLingua) As Boolean
   Dim i As Integer
   Dim strctls As String
   Dim Tempvalue
   
   On Error GoTo Errore
   LengFixTextRefresh = False
   RS_LabelText.MoveFirst
   For i = 1 To 20
      Cartellino.Fisso(i) = ""
   Next
   For i = 1 To 20
     ' If RS_LabelText.EOF = False Then Exit Function
      If Left(RS_LabelText("TagName"), 3) = "MAN" Then
         Tempvalue = Val(Right(RS_LabelText("TagName"), 2))
         Cartellino.Fisso(Val(Right(RS_LabelText("TagName"), 2))) = RS_LabelText(Cartellino.ColonnaLingua(inLingua))
      Else
         Tempvalue = 10 + NumeroAutoVisibile(RS_LabelText("TagName"))
         Cartellino.Fisso(10 + NumeroAutoVisibile(RS_LabelText("TagName"))) = RS_LabelText(Cartellino.ColonnaLingua(inLingua))
      End If
      RS_LabelText.MoveNext
   Next
   LengFixTextRefresh = True
   Exit Function
Errore:
End Function

Function FixVisibleRefresh() As Boolean
   Dim i As Integer
   Dim strAPP As String
   
   On Error GoTo Errore
   FixVisibleRefresh = False
   '============ tutti invisibili ===========
   For i = 1 To 20
     Cartellino.FissoVisibile(i) = False
     'Cartellino.manIndice(i) = 0
     DatiCartellino(i) = 0
   Next
   '=========================================
   For i = 0 To ManCount
      strAPP = Trim(Replace(CTLS(i).Nome, Chr(0), Chr(32)))
      If UCase(Left(strAPP, 3)) = "MAN" Then
         Cartellino.FissoVisibile(Val(Right(strAPP, 2))) = CTLS(i).Visibile
         'Cartellino.manIndice(Val(Right(strAPP, 2))) = i
         DatiCartellino(Val(Right(strAPP, 2))) = i
      Else
         If strAPP <> "" Then
            Dim k
            
            k = NumeroAutoVisibile(strAPP) + 10
            Cartellino.FissoVisibile(NumeroAutoVisibile(strAPP) + 10) = CTLS(i).Visibile
            'Cartellino.manIndice(NumeroAutoVisibile(strAPP) + 10) = i
            DatiCartellino(NumeroAutoVisibile(strAPP) + 10) = i
         End If
      End If
   Next
   FixVisibleRefresh = True
   Exit Function
   
Errore:
End Function
Function NumeroAutoVisibile(ByVal inNome As String) As Integer
   Dim strctls As String
   
   On Error GoTo Errore
   NumeroAutoVisibile = 0
   strctls = Trim(Replace(inNome, Chr(0), Chr(32)))
   Select Case strctls
      Case "Date"
           NumeroAutoVisibile = 1
      Case "Hour"
           NumeroAutoVisibile = 2
      Case "Bundle no."
           NumeroAutoVisibile = 3
      Case "Pcs/Bundle"
           NumeroAutoVisibile = 4
      Case "Dimensions"
           NumeroAutoVisibile = 5
      Case "Length"
           NumeroAutoVisibile = 6
      Case "Thickness"
           NumeroAutoVisibile = 7
      Case "Weight"
           NumeroAutoVisibile = 8
      Case "OrderDescr"
           NumeroAutoVisibile = 9
      Case Else
           NumeroAutoVisibile = 10
   End Select
   Exit Function
   
Errore:
End Function

Function LoadTicketFile(ByVal inPercorso As String) As Boolean
   Dim Lung As Long
   Dim Strname As String
   Dim TempName As String
   Dim Lenfile As Long
   Dim Dat$
   Dim i As Integer
   
   On Error GoTo Errore
   
   LoadTicketFile = False
   ManCount = 0
   
   If inPercorso <> "" And FileEsistente(inPercorso) Then
      Lung = Len(Writetemp)
      Strname = inPercorso
      TempName = Replace(inPercorso, "MairTicket", "TempMairTicket")
      
      Open Strname For Binary As #1
      Lenfile = LOF(1)
      Dat$ = Input(Lenfile, #1)
      Close 1
      
      Dat$ = Right$(Dat$, Len(Dat$) - 2)
      For i = 1 To 20
          With Writetemp
            .FontBold = Left$(Dat$, Len(Writetemp.FontBold))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.FontBold))
            .FontItal = Left$(Dat$, Len(Writetemp.FontItal))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.FontItal))
            .FontName = Left$(Dat$, Len(Writetemp.FontName))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.FontName))
            .FontSize = Left$(Dat$, Len(Writetemp.FontSize))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.FontSize))
            .Nome = Left$(Dat$, Len(Writetemp.Nome))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.Nome))
            .PosX = Val(Left$(Dat$, Len(Writetemp.PosX)))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.PosX))
            .PosY = Val(Left$(Dat$, Len(Writetemp.PosY))) - 690
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.PosY))
            .Visibile = Left$(Dat$, Len(Writetemp.Visibile))
            Dat$ = Right$(Dat$, Len(Dat$) - Len(Writetemp.Visibile))
         End With
                  
'         On Error Resume Next
'         Unload Man(i)
         On Error GoTo Errore
         
         ' controlla se c'è il nome, quindi crea l'oggetto
         If Trim(Replace(Writetemp.Nome, Chr(0), Chr(32))) <> "" Then
            ManCount = ManCount + 1
'            On Error Resume Next
'            Load Man(i)
'            On Error GoTo 0
'            Man(i).Visible = CBool(Writetemp.Visibile)
'            Man(i).Move Writetemp.PosX, Writetemp.PosY
'            Man(i).Font.Name = Trim(Writetemp.FontName)
'            Man(i).Font.Bold = CBool(Writetemp.FontBold)
'            Man(i).Font.Italic = CBool(Writetemp.FontItal)
'            Man(i).Font.Size = Writetemp.FontSize
'            Man(i).ToolTipText = Trim(Writetemp.Nome)
'            If CBool(Writetemp.Visibile) Then
'               Man(i).ForeColor = vbRed
'            Else
'               Man(i).ForeColor = vbBlack
'            End If
'            Man(i) = TestoCTL(Writetemp.Nome)
'            Man(i).ZOrder
            With Writetemp
                CTLS(i).FontBold = CBool(.FontBold)
                CTLS(i).FontItal = CBool(.FontItal)
                CTLS(i).FontName = .FontName
                CTLS(i).FontSize = Val(.FontSize)
                CTLS(i).Nome = .Nome
                CTLS(i).PosX = Val(.PosX)
                CTLS(i).PosY = Val(.PosY)
                CTLS(i).Visibile = CBool(.Visibile)
         End With
         End If
      Next
      
      On Error Resume Next
      
      Open TempName For Binary As #1
      Put #1, , Dat$
      Close #1
      Do
        DoEvents
      Loop Until FileEsistente(TempName) = True

'      Open TempName For Input As 1
'      OLE1.ReadFromFile 1
'      Close 1
      
      Kill TempName
      
      Do
          DoEvents
      Loop Until FileEsistente(TempName) = False

   End If
   LoadTicketFile = True
   Exit Function
   
Errore:

End Function

Function TestoCTL(ByVal Nome As String) As String
      Select Case Trim(Nome)
      Case "Date"
           TestoCTL = "mm/dd/yyyy"
      Case "Hour"
           TestoCTL = "mm:ss"
      Case "Bundle no."
           TestoCTL = "#####"
      Case "Pcs/Bundle"
           TestoCTL = "#####"
      Case "Dimensions"
           TestoCTL = "Diam.[LxH]"
      Case "Length"
           TestoCTL = "Tube lenght"
      Case "Thickness"
           TestoCTL = "Tube[T]"
      Case "Weight"
           TestoCTL = "Weight"
      Case Else
           TestoCTL = Trim(Nome)
      End Select
End Function

Private Sub Command1_Click()
   Me.Hide
End Sub

Private Sub OLE1_Click()
   Me.Hide
End Sub

