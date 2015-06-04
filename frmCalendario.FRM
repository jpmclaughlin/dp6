VERSION 5.00
Begin VB.Form frmcalendar 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendar"
   ClientHeight    =   4290
   ClientLeft      =   930
   ClientTop       =   1485
   ClientWidth     =   8145
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4290
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   2685
      Left            =   5160
      TabIndex        =   39
      Top             =   30
      Width           =   2985
      Begin VB.ComboBox cboyear 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   2745
      End
      Begin VB.ComboBox cbomonth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1320
         Width           =   2745
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Date"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   1890
         Width           =   615
      End
      Begin VB.Label lbldate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   2130
         Width           =   2715
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Year"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Month"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1575
      Left            =   5160
      TabIndex        =   36
      Top             =   2730
      Width           =   2985
      Begin VB.CommandButton cmdok 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   210
         Width           =   1785
      End
      Begin VB.CommandButton cmdcancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   540
         TabIndex        =   37
         Top             =   870
         Width           =   1785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Day"
      Height          =   555
      Left            =   0
      TabIndex        =   34
      Top             =   30
      Width           =   5145
      Begin VB.Label lblday 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   315
         Left            =   90
         TabIndex        =   35
         Top             =   180
         Width           =   5025
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   5145
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2520
         TabIndex        =   33
         Top             =   2610
         Width           =   810
      End
      Begin VB.Label lblPrec 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   30
         TabIndex        =   32
         Top             =   330
         Width           =   810
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   30
         Left            =   1980
         TabIndex        =   31
         Top             =   2610
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   29
         Left            =   1440
         TabIndex        =   30
         Top             =   2610
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   28
         Left            =   900
         TabIndex        =   29
         Top             =   2610
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   27
         Left            =   4140
         TabIndex        =   28
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   26
         Left            =   3600
         TabIndex        =   27
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   25
         Left            =   3060
         TabIndex        =   26
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   24
         Left            =   2520
         TabIndex        =   25
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   23
         Left            =   1980
         TabIndex        =   24
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   22
         Left            =   1440
         TabIndex        =   23
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   21
         Left            =   900
         TabIndex        =   22
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   20
         Left            =   4140
         TabIndex        =   21
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   19
         Left            =   3600
         TabIndex        =   20
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   18
         Left            =   3060
         TabIndex        =   19
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   17
         Left            =   2520
         TabIndex        =   18
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   16
         Left            =   1980
         TabIndex        =   17
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   15
         Left            =   1440
         TabIndex        =   16
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   14
         Left            =   900
         TabIndex        =   15
         Top             =   1470
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   13
         Left            =   4140
         TabIndex        =   14
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   12
         Left            =   3600
         TabIndex        =   13
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   11
         Left            =   3060
         TabIndex        =   12
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   10
         Left            =   2520
         TabIndex        =   11
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   9
         Left            =   1980
         TabIndex        =   10
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   8
         Left            =   1440
         TabIndex        =   9
         Top             =   900
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   900
         TabIndex        =   8
         Top             =   870
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   6
         Left            =   4140
         TabIndex        =   7
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   5
         Left            =   3600
         TabIndex        =   6
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   4
         Left            =   3060
         TabIndex        =   5
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   3
         Left            =   2520
         TabIndex        =   4
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   2
         Left            =   1980
         TabIndex        =   3
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblnumber 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
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
         Height          =   465
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   330
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmcalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code has been developed for EVERYONE'S use
' don't re-distribute this without ALL original files!!
'Phil Jones 1994

Option Explicit
Dim selectedate%

Private Sub cbomonth_click()
Call setday
Call lblnumber_click(selectedate% - 1)
End Sub

Private Sub cboyear_Click()
Static once% ' get rid of first click event
If Not once Then
    once = True
    Exit Sub
End If
Call cbomonth_click

End Sub

Private Sub checkdate(month1%, year1%)
Dim i%, value%, date1$

For i% = 28 To 32
    date1$ = (Str$(month1%) + "/" + Str$(i%) + "/" + Str$(year1%))
        If IsDate(date1$) Then
            value% = i%
        Else
            Call displaynumbers(value%)
            Exit Sub
        End If
Next i%
End Sub

Private Sub cmdcancel_Click()
  Unload Me
End Sub

Private Sub cmdok_Click()
    Dim month1%, day1%, year1%, date1$
    
    day1% = selectedate%
    month1% = cbomonth.ListIndex + 1
    year1% = cboyear.ListIndex + 1960
    date1$ = (Str$(month1%) + "/" + Str$(day1%) + "/" + Str$(year1%))
    date1$ = Format$(date1$, "general date")
    Calendario.Day = day1%
    Calendario.Month = month1%
    Calendario.year = year1%
    Unload Me
End Sub

Private Function determinemonth%()
Dim i%
i% = cbomonth.ListIndex 'which month is selected?
determinemonth% = i% + 1
End Function

Private Function determineyear%()
Dim i%
i% = cboyear.ListIndex 'which year was selected?
If i% = -1 Then Exit Function 'problem!!
determineyear% = CInt(Trim(cboyear.List(i%)))
End Function

Private Sub displaynumbers(number%)
Dim i%
For i% = 28 To 30
    lblnumber(i%).Visible = False
Next i%
For i% = 28 To number% - 1
    lblnumber(i%).Visible = True
Next i%

End Sub

Private Sub fillcbomonth()
cbomonth.AddItem "January"
cbomonth.AddItem "February"
cbomonth.AddItem "March"
cbomonth.AddItem "April"
cbomonth.AddItem "May"
cbomonth.AddItem "June"
cbomonth.AddItem "July"
cbomonth.AddItem "August"
cbomonth.AddItem "September"
cbomonth.AddItem "October"
cbomonth.AddItem "November"
cbomonth.AddItem "December"


End Sub

Private Sub fillcboyear()
Dim i%
For i% = 1960 To 2060 'put whatever years tyou want here,
    cboyear.AddItem Str$(i%) 'but don't forget to also change the code in setdate
Next i%

End Sub

Private Sub Form_Load()

selectedate% = CInt(Format$(Now, "dd"))

'fill month combo box
Call fillcbomonth

'fill year combo box
Call fillcboyear

'put current date and year im combo box
Call setdate

'set current name for day
Dim r%, caption1$
r% = Weekday(Format$(CDate(Calendario.Month & "/" & Calendario.Day & "/" & Calendario.year), "general date"))
If r% = 1 Then
    caption1$ = "Sunday"
ElseIf r% = 2 Then
    caption1 = "Monday"
ElseIf r% = 3 Then
    caption1 = "Tuesday"
ElseIf r% = 4 Then
    caption1 = "Wednesday"
ElseIf r% = 5 Then
    caption1 = "Thursday"
ElseIf r% = 6 Then
    caption1 = "Friday"
Else
    caption1 = "Saturday"
End If
lblday.Caption = caption1$
lblnumber_click (Calendario.Day - 1)
End Sub

Private Sub Label2_Click()
   If cbomonth.ListIndex < 11 Then
      cbomonth.ListIndex = cbomonth.ListIndex + 1
      lblnumber_click (0)
   Else
      If cboyear.ListIndex < cboyear.ListCount - 1 Then
         cboyear.ListIndex = cboyear.ListIndex + 1
         cbomonth.ListIndex = 0
         lblnumber_click (0)
      End If
   End If
End Sub

Private Sub lblnumber_click(Index As Integer)
Dim i%
On Error GoTo err1
For i% = 0 To 30
    lblnumber(i%).BackColor = &HC0C0C0
Next i%
If lblnumber(Index).BackColor = vbRed Then
    lblnumber(Index).BackColor = &HC0C0C0
Else
    lblnumber(Index).BackColor = vbRed
End If
selectedate% = Index + 1
Dim month1%, day1%, year1%, date1$
day1% = selectedate%
month1% = cbomonth.ListIndex + 1
year1% = cboyear.ListIndex + 1960
date1$ = (Str$(month1%) + "/" + Str$(day1%) + "/" + Str$(year1%))
'date1$ = Format$(date1$, "general date")
Dim r%
Dim caption1$
r% = Weekday(date1$)
If r% = 1 Then
    caption1$ = "Sunday"
ElseIf r% = 2 Then
    caption1 = "Monday"
ElseIf r% = 3 Then
    caption1 = "Tuesday"
ElseIf r% = 4 Then
    caption1 = "Wednesday"
ElseIf r% = 5 Then
    caption1 = "Thursday"
ElseIf r% = 6 Then
    caption1 = "Friday"
Else
    caption1 = "Saturday"
End If
lblday.Caption = caption1$
lbldate.Caption = Format$(Format(date1$, "mm-dd-yyyy"), "long date")

err1:
    If Err = 0 Then Exit Sub
    If Err = 13 Then
        selectedate% = selectedate% - 1
    Exit Sub
    End If
    End Sub

Private Sub setdate()
'year
Dim r%, i%
r% = CInt(Calendario.year)
i% = r% - 1960
cboyear.ListIndex = i%

'month
r% = CInt(Calendario.Month)
cbomonth.ListIndex = (r% - 1)

'day
r% = CInt(Calendario.Day)
lblnumber(r% - 1).BorderStyle = 1
selectedate% = r%

End Sub

Private Sub setday()
Dim month1%, year1%
month1% = determinemonth()
year1% = determineyear()
Call checkdate(month1%, year1%)

End Sub

Private Sub lblPrec_Click()
   If cbomonth.ListIndex > 0 Then
      cbomonth.ListIndex = cbomonth.ListIndex - 1
      lblnumber_click (0)
   Else
      If cboyear.ListIndex > 0 Then
         cboyear.ListIndex = cboyear.ListIndex - 1
         cbomonth.ListIndex = 11
         lblnumber_click (0)
      End If
   End If
End Sub
