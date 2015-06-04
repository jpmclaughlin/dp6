VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmEditor 
   Caption         =   "Help editor"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6945
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser webPreview 
      Height          =   2775
      Left            =   4110
      TabIndex        =   1
      Top             =   150
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   4895
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtEditor 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Nuovo"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Apri"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Salva"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "Salva &come..."
         Index           =   3
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "E&sci"
         Index           =   7
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditItem 
         Caption         =   "Cu&t"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Paste"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Select &All"
         Index           =   6
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Visualizza"
      Begin VB.Menu mnuViewItem 
         Caption         =   "&Editor"
         Index           =   0
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewItem 
         Caption         =   "&Anteprima"
         Index           =   1
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Inserisci"
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Intest"
         Index           =   0
         Begin VB.Menu mnuInsertHeaderItem 
            Caption         =   "Intest &1"
            Index           =   0
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuInsertHeaderItem 
            Caption         =   "Intest &2"
            Index           =   1
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuInsertHeaderItem 
            Caption         =   "Intest &3"
            Index           =   2
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuInsertHeaderItem 
            Caption         =   "Intest &4"
            Index           =   3
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnuInsertHeaderItem 
            Caption         =   "Intest &5"
            Index           =   4
            Shortcut        =   ^{F5}
         End
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Font"
         Index           =   1
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Paragrafo"
         Index           =   2
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Bold"
         Index           =   3
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Italic"
         Index           =   4
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "Ta&g Pair"
         Index           =   5
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "I&mmagine"
         Index           =   6
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Commento"
         Index           =   7
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&Tabella"
         Index           =   8
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "Hyperlin&k"
         Index           =   9
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "F&orm e controlli"
         Index           =   11
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&Form"
            Index           =   0
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&TextBox"
            Index           =   1
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "Textbox &Area"
            Index           =   2
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&Password"
            Index           =   3
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&RadioButton"
            Index           =   4
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&CheckBox"
            Index           =   5
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&Select"
            Index           =   6
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "S&ubmit"
            Index           =   8
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "R&eset"
            Index           =   9
         End
         Begin VB.Menu mnuInsertFormItem 
            Caption         =   "&Generico"
            Index           =   10
         End
      End
      Begin VB.Menu mnuInsertItem 
         Caption         =   "&VBScript codice"
         Index           =   12
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The contents of a new page
Const HTML_NEWPAGE = "<HTML>§<HEAD>§" _
                    & "<TITLE>{Titolo della pagina}</TITLE>§" _
                    & "</HEAD>§§" _
                    & "<BODY>§§</BODY>§</HTML>§"
Const DEF_CAPTION = "Help Editor"

Dim IsDirty As Boolean
Dim FileName As String

Private Sub Form_Activate()
    Static initDone As Boolean
    
    If initDone Then Exit Sub
    initDone = True
    webPreview.GoHome
    
End Sub

Private Sub Form_Load()
   ' mnuFileItem_Click (0)
End Sub

Private Sub Form_Resize()
    txtEditor.Move 0, 0, ScaleWidth, ScaleHeight
    webPreview.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0            'nuovo
            If ConfirmCommand Then
                txtEditor.Text = ""
                InsertText HTML_NEWPAGE
                mnuViewItem_Click 0
                IsDirty = False
                FileName = ""
                Caption = DEF_CAPTION
            End If
        Case 1            ' apri
            If ConfirmCommand Then OpenFile
        Case 2            ' Salva
            SaveFile
        Case 3            ' Salva come
            SaveFile , True
        Case 5            ' esci
            Unload Me
    End Select
End Sub

Private Sub mnuEdit_Click()
    mnuEditItem(2).Enabled = (txtEditor.SelText <> "")
    mnuEditItem(3).Enabled = (txtEditor.SelText <> "")
    mnuEditItem(4).Enabled = (Clipboard.GetText <> "")
End Sub

Private Sub mnuEditItem_Click(Index As Integer)
    Select Case Index
        Case 0          ' undo
        Case 2          ' cut
            Clipboard.Clear
            Clipboard.SetText txtEditor.SelText
            txtEditor.SelText = ""
        Case 3          ' copy
            Clipboard.Clear
            Clipboard.SetText txtEditor.SelText
        Case 4          ' paste
            txtEditor.SelText = Clipboard.GetText
        Case 6          ' select all
            txtEditor.SelStart = 0
            txtEditor.SelLength = Len(txtEditor.Text)
    End Select
End Sub

Private Sub mnuInsert_Click()
    mnuViewItem_Click 0
End Sub

Private Sub mnuInsertItem_Click(Index As Integer)
    Dim Text As String, SurroundSelText As Boolean
    
    Select Case Index
        Case 0         '
        Case 1
            Text = "<FONT FACE=""Fontname"" SIZE=FontSize COLOR=""#RRGGBB"">{Text}</FONT>"
            SurroundSelText = True
        Case 2         '
            Text = "<P>§"
        Case 3         '
            Text = "<B>{Text}</B>"
            SurroundSelText = True
        Case 4         '
            Text = "<I>{Text}</I>"
            SurroundSelText = True
        Case 5         '
            Text = InputBox("Enter tag contents")
            If Len(Text) Then
                Text = UCase$(Text)
                Text = "<" & Text & ">{Text}</" & Text & ">"
                SurroundSelText = True
            End If
        Case 6         '
            Text = "<IMG SRC=""{filename}"">"
        Case 7         '
            Text = "<!-- {comment text}§-->§"
            SurroundSelText = True
        Case 8        '
            Dim frm As New frmTable
            frm.Show vbModal
            Text = frm.HTMLText
        Case 9
            Text = "<A HREF=""url"">{Text}</A>"
            SurroundSelText = True
        Case 12        '
            Text = "<SCRIPT LANGUAGE=""VBScript"">§{your code}§</SCRIPT>§"
    End Select
    InsertText Text, SurroundSelText
End Sub

Private Sub mnuInsertHeaderItem_Click(Index As Integer)
    Dim Text As String
    Text = Trim$(Index + 1)
    InsertText "<H" & Text & ">{Level " & Text & " Heading}</H" & Text & ">§"
End Sub

Private Sub mnuInsertFormItem_Click(Index As Integer)
    Dim Text As String
    Dim frm As New frmMultiple
    
    Select Case Index
        Case 0        '
            Text = "<FORM NAME=""formname"">§{}§</FORM>§"
        Case 1        '
            Text = "<INPUT TYPE=Text NAME=""{ControlName}"" VALUE="""">§"
        Case 2        '
            Text = "<TEXTAREA NAME=""{ControlName}"" ROWS=5 COLS=30></TEXTAREA>§"
        Case 3        '
            Text = "<INPUT TYPE=Password NAME=""{ControlName}"" VALUE="""">§"
        Case 4        '
            frm.ControlType = 1
            frm.Show vbModal
            InsertText frm.HTMLText
        Case 5        '
            Text = "<INPUT TYPE=Checkbox NAME={ControlName} CHECKED>§"
        Case 6        '
            frm.ControlType = 2
            frm.Show vbModal
            InsertText frm.HTMLText
        Case 8        '
            Text = "<INPUT TYPE=SUBMIT VALUE=""{Submit}"">§"
        Case 9        '
            Text = "<INPUT TYPE=RESET VALUE=""{Reset}"">§"
        Case 10
            Text = "<INPUT TYPE=BUTTON NAME=""ButtonName"" VALUE=""{Button Caption}"">§"
    End Select
    
    InsertText Text
        
End Sub



Private Sub InsertText(ByVal Text As String, Optional SurroundSelText As Boolean)
    Dim i As Long, j As Long, SelStart As Long
    
    If Len(Text) = 0 Then Exit Sub
    
    Text = Replace(Text, "§", vbCrLf)
    
    i = InStr(Text, "{")
    If i Then
        j = InStr(i, Text, "}")
        If j Then
            If SurroundSelText And txtEditor.SelLength > 0 Then
                Text = Left$(Text, i - 1) & txtEditor.SelText & Mid$(Text, j + 1)
                j = i + txtEditor.SelLength + 1
            Else
                Text = Left$(Text, i - 1) & Mid$(Text, i + 1, j - i - 1) & Mid$(Text, j + 1)
            End If
            SelStart = txtEditor.SelStart
        End If
    End If
    
    txtEditor.SelText = Text
    If i > 0 And j > 0 Then
        txtEditor.SelStart = SelStart + i - 1
        txtEditor.SelLength = j - i - 1
    End If
End Sub

Private Sub mnuViewItem_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0            ' editor
            If webPreview.Visible Then LoadIntoBrowser ""
            txtEditor.ZOrder
            txtEditor.SetFocus
            webPreview.Visible = False
            
            
        Case 1            ' anteprima
            webPreview.Visible = True
            webPreview.ZOrder
            LoadIntoBrowser txtEditor.Text
            webPreview.SetFocus
    End Select
End Sub

Private Sub LoadIntoBrowser(Text As String)
    
  '  On Error GoTo Error_Handler
    
    With webPreview.Document
        .Clear
        .Open
        .write Text
        .Close
    End With
Error_Handler:

End Sub

Private Sub txtEditor_Change()
    IsDirty = True
End Sub

Private Function ConfirmCommand() As Boolean
    If IsDirty Then
        Select Case MsgBox("This file has been changed? Do you want to save it?", vbYesNoCancel + vbExclamation)
            Case vbYes
                ConfirmCommand = SaveFile()
            Case vbNo
                ConfirmCommand = True
            Case vbCancel
                '
        End Select
    Else
        ConfirmCommand = True
    End If
End Function

Function SaveFile(Optional FilePath As String, Optional showDialog As Boolean) As Boolean
    
    On Error Resume Next
    
    If FilePath = "" Then FilePath = FileName
    
    If FilePath = "" Or showDialog Then
        With CommonDialog1
            .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly
            .FileName = FilePath
            .Filter = "HTML Files (*.htm;*.html)|*.htm;*.html|All Files|*.*"
            .FilterIndex = 1
            .CancelError = True
            .ShowSave
            If Err Then Exit Function
            FilePath = .FileName
        End With
    End If
    
    Open FilePath For Output As #1
    Print #1, , txtEditor.Text;
    Close #1
    
    If Err = 0 Then
        FileName = FilePath
        IsDirty = False
        SaveFile = True
        Caption = DEF_CAPTION & " - " & FileName
    End If
End Function

Function OpenFile(Optional FilePath As String, Optional showDialog As Boolean) As Boolean
    
    On Error Resume Next
    
    If FilePath = "" Or showDialog Then
        With CommonDialog1
            .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
            .FileName = FilePath
            .Filter = "HTML Files (*.htm;*.html)|*.htm;*.html|All Files|*.*"
            .FilterIndex = 1
            .CancelError = True
            .ShowOpen
            If Err Then Exit Function
            FilePath = .FileName
        End With
    End If
    
    Open FilePath For Input As #1
    txtEditor.Text = Input$(LOF(1), 1)
    Close #1
    
    If Err = 0 Then
        FileName = FilePath
        IsDirty = False
        OpenFile = True
        Caption = DEF_CAPTION & " - " & FileName
    End If
End Function

