VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form DialogSerials 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10695
   ClientLeft      =   2715
   ClientTop       =   3390
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   555
      Left            =   450
      Top             =   150
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   979
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   11250
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9540
      Width           =   2085
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9540
      Width           =   1965
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridFiletti 
      Bindings        =   "DialogFiletti.frx":0000
      Height          =   9135
      Left            =   150
      TabIndex        =   2
      Top             =   210
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   16113
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "Type"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "Step"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(1)._Alignment=   7
      _Band(0)._MapCol(2)._Name=   "Lenght"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(2)._Alignment=   7
      _Band(0)._MapCol(3)._Name=   "Speed"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(3)._Alignment=   7
      _Band(0)._MapCol(4)._Name=   "Bookmark"
      _Band(0)._MapCol(4)._RSIndex=   4
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1245
      Index           =   1
      Left            =   13470
      TabIndex        =   3
      Top             =   240
      Width           =   1245
      _extentx        =   2196
      _extenty        =   2196
      txttext         =   " "
      txttop          =   5
      txtleft         =   5
      imgtop          =   15
      imgleft         =   18
      icona           =   "..\Bitmap\Icone\ARW03UP.ICO"
      imgh            =   50
      imgw            =   50
      imgallarga      =   -1  'True
      btype           =   3
      tx              =   ""
      enab            =   -1  'True
      font            =   "DialogFiletti.frx":0015
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      fcol            =   0
   End
   Begin dp6.XPButton XPButton1 
      Height          =   1275
      Index           =   2
      Left            =   13500
      TabIndex        =   4
      Top             =   8130
      Width           =   1245
      _extentx        =   2196
      _extenty        =   2249
      txttext         =   " "
      txttop          =   5
      txtleft         =   5
      imgtop          =   15
      imgleft         =   18
      icona           =   "..\Bitmap\Icone\ARW03DN.ICO"
      imgh            =   50
      imgw            =   50
      imgallarga      =   -1  'True
      btype           =   3
      tx              =   ""
      enab            =   -1  'True
      font            =   "DialogFiletti.frx":0039
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   65535
      fcol            =   0
   End
End
Attribute VB_Name = "DialogSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public SerialNumber As String
Public GradeNumber As String
Public RecipeName As String
Private PosizioneAttuale As String

Private Sub CancelButton_Click()
   SerialNumber = ""
   GradeNumber = ""
   Me.Hide
End Sub

Private Sub Form_Load()
   Call RefreshFiletti
   SerialNumber = ""
   GradeNumber = ""
   RecipeName = ""
End Sub

Private Sub OKButton_Click()
        GridFiletti.Col = 5
        SerialNumber = GridFiletti.Text
        GridFiletti.Col = 2
        GradeNumber = GridFiletti.Text
        GridFiletti.Col = 1
        RecipeName = GridFiletti.Text
'  With FilettoForm.AdoFiletti
''        .Recordset.MoveFirst
''        .Recordset.Find ("Bookmark=1")
''        FilettoForm.CancBook = True
''        .Recordset("Bookmark") = 0
''        .Recordset.AbsolutePosition = GridFiletti.RowSel
''        FilettoForm.CancBook = False
''        .Recordset("Bookmark") = 1
''        .Recordset.Update
'    End With
    
'   Set FilettoForm.AdoFiletti.Recordset.ActiveConnection = Nothing
'   FilettoForm.Command1.Enabled = True
'   FilettoForm.Command1.BackColor = &HFFFF&
   Me.Hide
End Sub

Public Sub RefreshFiletti()
    With Adodc1
        .CommandType = adCmdText
        .ConnectionString = Connessione.ConnectionString
        .RecordSource = "SELECT * FROM Q_itemscode"
        .Refresh
             
        'GridFiletti.ColWidth(0) = 0
        GridFiletti.ColWidth(0) = 0
        GridFiletti.ColWidth(1) = 3000
        GridFiletti.ColWidth(2) = 5000
        GridFiletti.ColWidth(3) = 0
        GridFiletti.ColWidth(4) = 1800
        GridFiletti.ColWidth(5) = 3000
             
        GridFiletti.Row = 1
'        GridFiletti.Row = .Recordset.AbsolutePosition
'        GridFiletti.Col = 0
'        GridFiletti.CellBackColor = &HFFFF00
'        GridFiletti.Col = 1
'        GridFiletti.CellBackColor = &HFFFF00
'        GridFiletti.Col = 2
'        GridFiletti.CellBackColor = &HFFFF00
'        GridFiletti.Col = 3
'        GridFiletti.CellBackColor = &HFFFF00
        
'        ' recupera informazioni circa il filetto selezionato
'        On Error GoTo MancaBookmark
'        GridFiletti.Row = .Recordset.AbsolutePosition
    End With
    Exit Sub

MancaBookmark:
        
        MsgBox "Manca il bit bookmark"

End Sub


Private Sub XPButton1_Click(Index As Integer)
   Select Case Index
   Case 1
            PosizioneAttuale = GridFiletti.TopRow
            PosizioneAttuale = PosizioneAttuale - 10
            If PosizioneAttuale < 0 Then PosizioneAttuale = 0
            If GridFiletti.Rows > 0 Then GridFiletti.TopRow = PosizioneAttuale
   Case 2
            PosizioneAttuale = GridFiletti.TopRow
            PosizioneAttuale = PosizioneAttuale + 10
            If PosizioneAttuale >= GridFiletti.Rows Then PosizioneAttuale = GridFiletti.Rows - 1
            If GridFiletti.Rows > 0 Then GridFiletti.TopRow = PosizioneAttuale
   End Select
End Sub
