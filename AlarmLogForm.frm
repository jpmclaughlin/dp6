VERSION 5.00
Begin VB.Form AlarmLogForm 
   Caption         =   "Alarm log"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   8775
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "AlarmLogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ricordarsi di chiudere e riaprire gli oggetti *Data
' qui utilizzati nella funzione "CompactDatabase"
' nel modulo "AlarmForm"



Private Sub Form_Load()
'    AlarmLogData.DataBaseName = AlarmForm.AlarmDataBaseName
'    AlarmLogDBGrid.Columns(0).Caption = Param.Text("DateTime")
'    AlarmLogDBGrid.Columns(1).Caption = Param.Text("State")
'    AlarmLogDBGrid.Columns(2).Caption = Param.Text("Number")
'    AlarmLogDBGrid.Columns(3).Caption = Param.Text("Text")
End Sub

'Private Sub PrintCommand_Click()
'    'PrinterDialog.ShowPrinter
'    ' printer. ...
'    ' Printer.EndDoc
'End Sub
'

'****************************************************
' Funzione per impedire la chiusura del form da parte
' dell'utente
'****************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormControlMenu Then
        On Error Resume Next
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

