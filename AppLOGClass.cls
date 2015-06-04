VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppLOGClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const NomeFileLog = "\AppLOG.txt"
Const EventiMax = 20000
Public NumEventi As Long
Public CancellaEventi As Boolean
Private ClassInited As Boolean

Function logEvent(ByVal evento As String) As Boolean
    Dim modo As Boolean
    Dim f As Integer
    
    On Error GoTo Errore
    logEvent = False
    f = FreeFile
    'verifica se cancellare eventi
    modo = False
    modo = FileEsistente(App.path & NomeFileLog)
    If NumEventi < EventiMax Then modo = True: NumEventi = 0
    'scrittura eventi
    If modo And CancellaEventi = False And ClassInited = True Then
       Open App.path & NomeFileLog For Append As #f
    Else
       Open App.path & NomeFileLog For Output As #f
    End If
    Print #f, Format(Time, "hh:mm:ss") & "|" & evento
    NumEventi = NumEventi + 1
    logEvent = True
    ClassInited = True
Errore:
    On Error Resume Next
    Close #f
End Function

Public Function logError(ByVal inPosiz As String) As Boolean
    Dim modo As Boolean
    Dim f As Integer
    
    On Error GoTo Errore
    logError = False
    f = FreeFile
    'verifica se cancellare eventi
    modo = False
    modo = FileEsistente(App.path & NomeFileLog)
    If NumEventi < EventiMax Then modo = True: NumEventi = 0
    'scrittura eventi
    If modo And CancellaEventi = False And ClassInited = True Then
       Open App.path & NomeFileLog For Append As #f
    Else
       Open App.path & NomeFileLog For Output As #f
    End If
    Print #f, Format(Time, "hh:mm:ss") & "|" & inPosiz & ":ERR=" & Err.LastDllError & " | " & Err.Description
    NumEventi = NumEventi + 1
    logError = True
    ClassInited = True
Errore:
    On Error Resume Next
    Close #f
End Function

Private Sub Class_Initialize()
   ClassInited = False
End Sub
