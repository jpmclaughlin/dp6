VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrdiniEvasiClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public IDOrdine As Integer
Public IDRicetta As String
Public DescrizioneOrdine As String
Public PresetPacchi As Integer
Public ModoCambioOrdine As Integer
Public LinguaCartellino As Integer
Public Visualizzato As Boolean

' campi ad aggiornamento manuale del cartellino
Private PrivPrinterData(10) As String

'************************************************************
' Funzioni di accesso ad dati privati
'************************************************************
Property Let CampoManuale(Index As Integer, value As String)
    If Index >= 1 And Index <= 10 Then
        PrivPrinterData(Index) = value
    End If
End Property

Property Get CampoManuale(Index As Integer) As String
    If Index >= 1 And Index <= 10 Then
        CampoManuale = PrivPrinterData(Index)
    Else
        CampoManuale = ""
    End If
End Property

Public Function Cancella(ByVal IDRic As String) As Boolean
    'Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
   ' On Error GoTo errore
    Cancella = False
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE ID=" & IDRic, Connessione, , adLockOptimistic, adCmdText
     
          If .EOF = False Then
             .Delete adAffectCurrent
             Cancella = True
          End If
          .Close
          .ActiveConnection = Nothing
    End With
    Set rs = Nothing
    'Set cn = Nothing
Exit Function

Errore:
MsgBox "Errore di connessione al database ordini", vbExclamation, "DATAPACK 6.0"
End Function
    
Public Function IndiceDuplicato(ByVal IDord As Integer) As Boolean
    'Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
      
  '  On Error GoTo errore
    IndiceDuplicato = False
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE ID=" & IDord, Connessione, , adLockReadOnly, adCmdText
     
          If .EOF = False Then IndiceDuplicato = True
          .Close
          .ActiveConnection = Nothing
    End With
    Set rs = Nothing
    'Set cn = Nothing
Exit Function

Errore:
MsgBox "Errore di connessione al database ordini", vbExclamation, "DATAPACK 6.0"
End Function
Public Sub DownloadData(ByVal IDord As Integer)
  '  Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
   ' cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE ID=" & IDord, Connessione, , adLockOptimistic, adCmdText
    
        If .EOF = False Then
           On Error Resume Next
                .Fields("Descrizione") = DescrizioneOrdine
                .Fields("IDRicetta") = IDRicetta
                .Fields("Visualizzato") = Visualizzato
                .Fields("GestioneAFineOrdine") = ModoCambioOrdine
                .Fields("NumPacchi") = PresetPacchi
                .Fields("LinguaCartellino") = LinguaCartellino
'                If Cliente <> "" Then
'                   .Fields("Cliente") = Cliente
'                Else
'                   .Fields("Cliente") = "---"
'                End If
'                .Fields("Tappatrice") = Tappatrice
'                 If Confez <> "" Then
'                   .Fields("Confez") = Confez
'                Else
'                   .Fields("Confez") = "0"
'                End If
'                If Lamiflex <> "" Then
'                   .Fields("Lamiflex") = Lamiflex
'                Else
'                   .Fields("Lamiflex") = "---"
'                End If
                For i = 1 To 10
                    .Fields("Cartellino" & i) = CampoManuale(i)
                Next i
                .Update
            On Error GoTo 0
        End If
        .Close
        .ActiveConnection = Nothing
    End With
    Set rs = Nothing
   ' Set cn = Nothing
End Sub
Public Function UploadData(ByVal IDord As Integer) As Boolean
    Dim i As Integer
   ' Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Errore
   ' cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE ID=" & IDord, Connessione, , adLockReadOnly, adCmdText
     
        If .EOF = False Then
            'On Error Resume Next
                DescrizioneOrdine = .Fields("Descrizione")
                IDRicetta = .Fields("IDRicetta")
                Visualizzato = .Fields("Visualizzato")
                ModoCambioOrdine = .Fields("GestioneAFineOrdine")
                PresetPacchi = .Fields("NumPacchi")
                LinguaCartellino = .Fields("LinguaCartellino")
'                Cliente = .Fields("Cliente")
'                Tappatrice = .Fields("Tappatrice")
'                Confez = .Fields("Confez")
'                Lamiflex = .Fields("Lamiflex")
                For i = 1 To 10
                    CampoManuale(i) = .Fields("Cartellino" & i)
                Next i
            On Error GoTo 0
            UploadData = True
        Else
            UploadData = False
        End If
        .Close
        .ActiveConnection = Nothing
    End With
    Set rs = Nothing
   ' Set cn = Nothing
    Exit Function

Errore:
   'MsgBox "Errore di connessione al database ordini", vbExclamation, "DATAPACK 6.0"
End Function


