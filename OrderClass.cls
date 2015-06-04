VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IDOrdine As Integer

' dati produzione
Public IDRicetta As String
' dati etichetta
Public Visualizzato As Boolean
Public Descrizione As String
Public PresetPacchi As Integer
Public ModoCambioOrdine As Integer
Public LinguaCartellino As Integer
Public TicketUnit As Boolean

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
    
    On Error GoTo Errore
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
      
    On Error GoTo Errore
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
    'Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE ID=" & IDord, Connessione, , adLockOptimistic, adCmdText
    
        If .EOF = False Then
           On Error Resume Next
                .Fields("Descrizione") = Descrizione
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
                .Fields("Unit") = TicketUnit
                For i = 1 To 10
                    .Fields("Cartellino" & i) = CampoManuale(i)
                Next i
                .Update
        End If
        .Close
        .ActiveConnection = Nothing
    End With
    Set rs = Nothing
    'Set cn = Nothing
End Sub
Public Function UploadData(ByVal IDord As Integer) As Boolean
    Dim i As Integer
    'Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Errore
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE ID=" & IDord, Connessione, , adLockReadOnly, adCmdText
     
        If .EOF = False Then
            'On Error Resume Next
                Descrizione = .Fields("Descrizione")
                IDRicetta = .Fields("IDRicetta")
                Visualizzato = .Fields("Visualizzato")
                ModoCambioOrdine = .Fields("GestioneAFineOrdine")
                PresetPacchi = .Fields("NumPacchi")
                LinguaCartellino = .Fields("LinguaCartellino")
'                Cliente = .Fields("Cliente")
'                Tappatrice = .Fields("Tappatrice")
'                Confez = .Fields("Confez")
'                Lamiflex = .Fields("Lamiflex")
                TicketUnit = .Fields("Unit")
                For i = 1 To 10
                    CampoManuale(i) = .Fields("Cartellino" & i)
                Next i
            UploadData = True
        Else
            UploadData = False
        End If
        .Close
        .ActiveConnection = Nothing
    End With
    Set rs = Nothing
    'Set cn = Nothing
    Exit Function

Errore:
   'MsgBox "Errore di connessione al database ordini", vbExclamation, "DATAPACK 6.0"
End Function

Public Function OrdineFuturo() As Long
    'Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
      
    On Error GoTo Errore
    OrdineFuturo = 0
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Orders WHERE Visualizzato=True", Connessione, , adLockReadOnly, adCmdText
     
         If .EOF = False Then OrdineFuturo = .Fields("ID")
         .Close
         .ActiveConnection = Nothing
    End With
    Set rs = Nothing
    'Set cn = Nothing
Exit Function

Errore:
MsgBox "Errore di connessione al database ordini", vbExclamation, "DATAPACK 6.0"
End Function
Public Function Client_OPEN_DBproduzione(Optional ByVal Client As Boolean) As Boolean
 
    Client_OPEN_DBproduzione = False
    On Error GoTo Errore
    Set Connessione = New ADODB.Connection
    Connessione.CursorLocation = adUseClient
    Connessione.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    Client_OPEN_DBproduzione = True
    Exit Function
Errore:
End Function

Public Sub Client_CLOSE_RSproduzione()
   On Error Resume Next
   RS_produzione.Close
   Set RS_produzione = Nothing
End Sub
Public Sub Client_CLOSE_DBproduzione()
   On Error Resume Next
   Set Connessione = Nothing
End Sub

Public Function Client_OrdiniInMacchina_CreateRS() As Boolean
   On Error GoTo Errore
   Client_OrdiniInMacchina_CreateRS = False
   Set RS_produzione = New ADODB.Recordset
   Client_OrdiniInMacchina_CreateRS = True
Errore:
End Function

Public Function Client_OrdiniInMacchina_Refresh(Optional ByVal Client As Boolean) As Boolean
    Dim rsstr As String
    
    'On Error Resume Next
    ' QUERY: data upload from Orders JOIN Recipes databases
    rsstr = "SELECT Orders.ID, Orders.Descrizione, Recipes.TipoTubo, Recipes.Altezza, Recipes.Larghezza, Recipes.Spessore, Recipes.Lunghezza, Orders.IDRicetta FROM Orders INNER JOIN Recipes ON Orders.IDRicetta = Recipes.ID ORDER BY Orders.ID"
    
    If Client Then
       If RS_produzione.State = adStateOpen Then RS_produzione.Close
       Connessione.CursorLocation = adUseClient
    End If
    Client_OrdiniInMacchina_Refresh = False
    With RS_produzione
         ' open the connection to the DB and refresh the data
         .Open rsstr, Connessione, , adLockReadOnly, adCmdText
         If .EOF = False Then Client_OrdiniInMacchina_Refresh = True
         If Client Then .ActiveConnection = Nothing ' close the connection to the DB
    End With
    Exit Function

Errore:
    MsgBox "DB 'Ordini' connection error", vbExclamation, "DATAPACK 6.0"
End Function

Public Function Client_Ordine_find(ByVal OrderNumber As String) As Boolean
    Dim rsstr As String
    
    On Error Resume Next
    Client_Ordine_find = False
    RS_produzione.MoveFirst
    With RS_produzione
       .Find "ID=" & CLng(OrderNumber), , adSearchForward, 1
       If .EOF = False Then Client_Ordine_find = True
    End With
    
Errore:
    
End Function

