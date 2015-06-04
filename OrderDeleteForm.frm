VERSION 5.00
Begin VB.Form OrderDeleteForm 
   BackColor       =   &H0000FFFF&
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OkCommand 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton CancelCommand 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancella ordine :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   300
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sei sicuro ?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   4035
   End
   Begin VB.Label OrderLabel 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "OrderDeleteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Public RecipeOrOrder As Boolean

Private Sub CancelCommand_Click()
    RecipeOrOrder = False
    OK = False
    Me.Hide
End Sub

Private Sub Form_Activate()
   OK = False
   ScritteMultilingua
End Sub

Private Sub OkCommand_Click()
    RecipeOrOrder = False
    OK = True
    Me.Hide
End Sub

Sub ScritteMultilingua()
    If RecipeOrOrder Then
       Label1.Caption = Param.Text("CancellaRicetta") & " : "
    Else
       Label1.Caption = Param.Text("Cancella ordine") & " : "
    End If
    Label2.Caption = Param.Text("Sei sicuro") & " ? "
    CancelCommand.Caption = Param.Text("Annulla")
End Sub
