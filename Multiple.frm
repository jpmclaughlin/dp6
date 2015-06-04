VERSION 5.00
Begin VB.Form frmMultiple 
   Caption         =   ".."
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
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
   ScaleHeight     =   3135
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkMultiple 
      Caption         =   "Selez. multipla"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtRows 
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Text            =   "1"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtOptions 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblRows 
      Caption         =   "Righe"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Opzioni"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nome controllo"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 1=radio button, 2=select
Public ControlType As Integer

Dim Text As String

Private Sub cmdOK_Click()
    Dim options() As String, i As Integer
    
    On Error GoTo Error_Handler
    
    If txtName = "" Or txtOptions = "" Then GoTo Error_Handler
    
    options = Split(txtOptions.Text, vbCrLf)
    
    If ControlType = 1 Then
        For i = 0 To UBound(options)
            Text = Text & "<INPUT TYPE=Radio NAME=""" & txtName & """" & IIf(i = 0, " CHECKED", "") & ">" & options(i) & "<BR>§"
        Next
    ElseIf ControlType = 2 Then
        Text = "<SELECT NAME=""" & txtName & """ SIZE=" & txtRows & IIf(chkMultiple, " MULTIPLE", "") & ">§"
        For i = 0 To UBound(options)
            Text = Text & "    <OPTION VALUE=" & (i + 1) & ">" & options(i) & "</OPTION>§"
        Next
        Text = Text & "</SELECT>§"
    End If

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

Private Sub Form_Load()
    If ControlType = 1 Then
        Caption = "RadioButton control"
        lblRows.Visible = False
        txtRows.Visible = False
        chkMultiple.Visible = False
    ElseIf ControlType = 2 Then
        Caption = "Select control"
    End If
End Sub
