VERSION 5.00
Begin VB.Form frmTable 
   Caption         =   "Tabella Html"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBorder 
      Alignment       =   1  'Right Justify
      Caption         =   "Bordo"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtColumns 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2520
      TabIndex        =   5
      Text            =   "2"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtRows 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2520
      TabIndex        =   3
      Text            =   "2"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtHeadRows 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero colonne "
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero righe normali: "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Numero righe intestazione: "
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Text As String

Private Sub cmdOK_Click()
    Dim headerRows As Integer, rows As Integer, columns As Integer
    Dim r As Integer, c As Integer
    
    On Error GoTo Error_Handler
    
    headerRows = CInt(txtHeadRows)
    rows = CInt(txtRows)
    columns = CInt(txtColumns)
    If headerRows < 0 Or rows < 0 Or headerRows + rows < 1 Or columns < 1 Then GoTo Error_Handler
    
    Text = "<TABLE" & IIf(chkBorder, " BORDER", "") & ">§"
    For r = 1 To headerRows
        Text = Text & "<TR>§"
        For c = 1 To columns
            Text = Text & "   <TH> HeadRow " & r & ", Column " & c & "</TH>§"
        Next
        Text = Text & "</TR>§"
    Next
    For r = 1 To rows
        Text = Text & "<TR>§"
        For c = 1 To columns
            Text = Text & "   <TD> HeadRow " & r & ", Column " & c & "</TD>§"
        Next
        Text = Text & "</TR>§"
    Next
    Text = Text & "</TABLE>§"

    Unload Me
    Exit Sub
    
Error_Handler:
    MsgBox "Invalid values", vbCritical

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Property Get HTMLText() As String
    HTMLText = Text
End Property


