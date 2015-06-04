VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Private Const FILE_ATTRIBUTE_ALL As Long = FILE_ATTRIBUTE_READONLY Or _
                                           FILE_ATTRIBUTE_HIDDEN Or _
                                           FILE_ATTRIBUTE_SYSTEM Or _
                                           FILE_ATTRIBUTE_ARCHIVE Or _
                                           FILE_ATTRIBUTE_NORMAL

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function GetFileAttributes Lib "kernel32" _
   Alias "GetFileAttributesA" _
  (ByVal lpFileName As String) As Long

Private Declare Function lstrlenW Lib "kernel32" _
  (ByVal lpString As Long) As Long
         
Private Declare Function CreateDirectory Lib "kernel32" _
   Alias "CreateDirectoryA" _
  (ByVal lpPathName As String, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Declare Function CopyFile Lib "kernel32" _
   Alias "CopyFileA" _
  (ByVal lpExistingFileName As String, _
   ByVal lpNewFileName As String, _
   ByVal bFailIfExists As Long) As Long


Private Declare Function LockWindowUpdate Lib "user32" _
  (ByVal hwndLock As Long) As Long

         
Private arrayFile() As String
         
Public Enum tCoerenza
       Uguale = 0
       Diversa = 1
       NonPresente = 2
       errImprevisto = 3
End Enum

Public Enum tErr
       FileOK = 0
       Sconosciuto = 1
       NoAutoriz = 2
       NonTrovato = 3
End Enum

Public DatidaScomporre As String
Public MaxCountFiles As Long

Private Enum DataLen
      lenAltezzapacco = 13
      lenDescrizione = 50
      lenFilaContros = 5
      lenGestFineOrdine = 2
      lenIDricetta = 50
      lenLargBase = 13
      lenLargEsa = 13
      lenLargMax = 13
      lenNumeroFile = 5
      lenNumeroTubi = 5
      lenNumPacchi = 5
      lenPesoPacco = 13
      lenRegge = 2
      lenTAltezza = 13
      lenTipoPacco = 1
      lenTipotubo = 1
      lenTLarghezza = 13
      lenTLunghezza = 13
      lenTPeso = 13
      lenTSpessore = 13
      lenVelMagneti = 3
      lenReggia = 13
      lenTubiFila = 5
End Enum

Private Type Formatostr
     IDRicetta As String * 50
     Descrizione As String * 50
     NumPacchi As Long
     GestFineOrdine As Long
     TipoTubo As Long
     TAltezza As Single
     TLarghezza As Single
     TLunghezza As Single
     TSpessore As Single
     TPeso As Single
     TipoPacco As Long
     NumeroTubi As Long
     NumeroFile As Long
     LargBase As Single
     LargMax As Single
     LargEsa As Single
     AltezzaPacco As Single
     FilaContros As Long
     PesoPacco As Single
     TubiFila(1 To 50) As Long
     Regge As Long
     Reggia(1 To 12) As Single
     VelMagneti As Long
End Type

Private DatiFromFile As Formatostr
Property Let NomeRicetta(ByVal inNuovo As String)
   DatiFromFile.IDRicetta = inNuovo
End Property
Property Get NomeRicetta() As String
   NomeRicetta = DatiFromFile.IDRicetta
End Property
Property Get Nomefiletrovato(ByVal Index As Long) As String
   If Index <= MaxCountFiles Then
      Nomefiletrovato = arrayFile(Index)
   End If
End Property
Property Get LungFile() As Long
   LungFile = 0
   LungFile = LungFile + DataLen.lenAltezzapacco + DataLen.lenDescrizione
   LungFile = LungFile + DataLen.lenFilaContros + DataLen.lenGestFineOrdine
   LungFile = LungFile + DataLen.lenIDricetta + DataLen.lenLargBase
   LungFile = LungFile + DataLen.lenLargEsa + DataLen.lenLargMax
   LungFile = LungFile + DataLen.lenNumeroFile + DataLen.lenNumeroTubi
   LungFile = LungFile + DataLen.lenNumPacchi + DataLen.lenPesoPacco
   LungFile = LungFile + DataLen.lenRegge + DataLen.lenTAltezza
   LungFile = LungFile + DataLen.lenTipoPacco + DataLen.lenTipotubo
   LungFile = LungFile + DataLen.lenTLarghezza + DataLen.lenTLunghezza
   LungFile = LungFile + DataLen.lenTPeso + DataLen.lenTSpessore + DataLen.lenVelMagneti
   LungFile = LungFile + DataLen.lenReggia * 12 + DataLen.lenTubiFila * 50
End Property
Sub initDati()
   Dim i As Integer

   With DatiFromFile
      .AltezzaPacco = 0
      .Descrizione = String(50, ".")
      .FilaContros = 0
      .GestFineOrdine = 0
      .IDRicetta = String(50, ".")
      .LargBase = 0
      .LargEsa = 0
      .LargMax = 0
      .NumeroFile = 0
      .NumeroTubi = 0
      .NumPacchi = 0
      .PesoPacco = 0
      .Regge = 0
      .TAltezza = 0
      .TipoPacco = 0
      .TipoTubo = 0
      .TLarghezza = 0
      .TLunghezza = 0
      .TPeso = 0
      .TSpessore = 0
      .VelMagneti = 0
      For i = 1 To 50
         If i < 13 Then .Reggia(i) = 0
         .TubiFila(i) = 0
      Next
   End With
End Sub
Function CompStringa() As String
   Dim i As Integer
   Dim a
   
   On Error GoTo errore
   CompStringa = ""
   With DatiFromFile
      CompStringa = CompStringa & CStr(Format(.AltezzaPacco, "000.000000000"))
      CompStringa = CompStringa & .Descrizione
      CompStringa = CompStringa & CStr(Format(.FilaContros, "00000"))
      CompStringa = CompStringa & CStr(Format(.GestFineOrdine, "00"))
      CompStringa = CompStringa & .IDRicetta
      CompStringa = CompStringa & CStr(Format(.LargBase, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.LargEsa, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.LargMax, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.NumeroFile, "00000"))
      CompStringa = CompStringa & CStr(Format(.NumeroTubi, "00000"))
      CompStringa = CompStringa & CStr(Format(.NumPacchi, "00000"))
      CompStringa = CompStringa & CStr(Format(.PesoPacco, "000000.000000"))
      CompStringa = CompStringa & CStr(Format(.Regge, "00"))
      CompStringa = CompStringa & CStr(Format(.TAltezza, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.TipoPacco, "0"))
      CompStringa = CompStringa & CStr(Format(.TipoTubo, "0"))
      CompStringa = CompStringa & CStr(Format(.TLarghezza, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.TLunghezza, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.TPeso, "000000.000000"))
      CompStringa = CompStringa & CStr(Format(.TSpessore, "000.000000000"))
      CompStringa = CompStringa & CStr(Format(.VelMagneti, "000"))
      For i = 1 To 50
         If i < 13 Then CompStringa = CompStringa & CStr(Format(.Reggia(i), "000.000000000"))
         CompStringa = CompStringa & CStr(Format(.TubiFila(i), "00000"))
      Next
   End With
   
errore:
End Function
Function TestFileLength(ByVal inNomefile As String, ByVal inLung As Long, ByRef outTestOK As Boolean) As tErr
    Dim errFile As tErr
    
    On Error GoTo errore
    outTestOK = False
    TestFileLength = tErr.FileOK
    errFile = LeggeFileDati(inNomefile)
    outTestOK = (Len(DatidaScomporre) >= inLung)
errore:
    TestFileLength = errFile
End Function

Function GestioneErrori(ByVal inNomefile As String) As Boolean
    On Error GoTo errore
    GestioneErrori = False
    If CancellaFile(inNomefile) = False Then GoTo errore
    GestioneErrori = True
errore:
End Function

Function CancellaFile(ByVal inNomefile As String) As Boolean
    On Error GoTo errore
    CancellaFile = False
    SetAttr inNomefile, vbArchive
    Kill inNomefile
    CancellaFile = True
errore:
End Function

Function LeggeFileDati(ByVal inFilenome As String) As tErr
   Dim f As Integer
   
   On Error GoTo errore
   Err.Clear
   LeggeFileDati = tErr.FileOK
   f = FreeFile()
   If (Dir(inFilenome, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "") = False Then
      LeggeFileDati = tErr.NonTrovato
      Exit Function
   End If
   Open inFilenome For Input As #f
   DatidaScomporre = Input(LOF(f), #f)
errore:
   On Error Resume Next
   Close #f
   LeggeFileDati = tErr.Sconosciuto
   Select Case Err.number
   Case 0
       DatidaScomporre = Replace(DatidaScomporre, vbCr, "")
       DatidaScomporre = Replace(DatidaScomporre, vbCrLf, "")
       DatidaScomporre = Replace(DatidaScomporre, vbLf, "")
       DatidaScomporre = Replace(DatidaScomporre, vbNullChar, "")
       LeggeFileDati = FileOK
   Case 70
       LeggeFileDati = NoAutoriz
   End Select
End Function
Function ScriveFileDati(ByVal inFilenome As String, inStringa As String, Optional inW_A As Boolean = False) As Boolean
   'Dim a$
   Dim f As Integer
   
   ScriveFileDati = False
   On Error GoTo errore
   f = FreeFile()
   'a$ = CompStringa
   inStringa = Trim(inStringa)
   If inStringa = "" Then Exit Function
   If inW_A Then
      Open inFilenome For Append Lock Read Write As #f
   Else
      Open inFilenome For Output Lock Read Write As #f
   End If
   Print #f, inStringa
   ScriveFileDati = True
errore:
   On Error Resume Next
   Close #f
End Function

Function CoerenzaRicetta() As tCoerenza
   Dim rs_ricerca As New ADODB.Recordset
   Dim tempOK As Boolean
   
   On Error GoTo errore
   CoerenzaRicetta = NonPresente
   rs_ricerca.CursorLocation = adUseClient
   rs_ricerca.Open "SELECT * FROM Ricette WHERE ID='" & Trim(DatiFromFile.IDRicetta) & "'", Connessione, adOpenKeyset, adLockReadOnly
   If rs_ricerca.EOF = False Then
      CoerenzaRicetta = Diversa
      tempOK = False
      tempOK = tempOK Or (CLng(rs_ricerca("TipoTubo")) <> DatiFromFile.TipoTubo)
      tempOK = tempOK Or (CSng(rs_ricerca("Altezza")) <> DatiFromFile.TAltezza)
      tempOK = tempOK Or (CSng(rs_ricerca("Larghezza")) <> DatiFromFile.TLarghezza)
      tempOK = tempOK Or (CSng(rs_ricerca("Lunghezza")) <> DatiFromFile.TLunghezza)
      tempOK = tempOK Or (CSng(rs_ricerca("Spessore")) <> DatiFromFile.TSpessore)
      tempOK = tempOK Or (CLng(rs_ricerca("TipoPacco")) <> DatiFromFile.TipoPacco)
      tempOK = tempOK Or (CLng(rs_ricerca("NumeroTubi")) <> DatiFromFile.NumeroTubi)
      tempOK = tempOK Or (CLng(rs_ricerca("NumeroFile")) <> DatiFromFile.NumeroFile)
      tempOK = tempOK Or (CLng(rs_ricerca("NumeroRegge")) <> DatiFromFile.Regge)
      If tempOK = False Then CoerenzaRicetta = Uguale
   End If
   rs_ricerca.Close
   Set rs_ricerca = Nothing
   Exit Function
errore:
   Set rs_ricerca = Nothing
   CoerenzaRicetta = tCoerenza.errImprevisto
End Function
Function FineLista(ByVal inMaxOrdini As Integer) As Boolean
   Dim rs_ricerca As New ADODB.Recordset
   
   On Error GoTo errore
   FineLista = False
   rs_ricerca.CursorLocation = adUseClient
   rs_ricerca.Open "SELECT Ordini.ID, Ordini.Visualizzato From Ordini WHERE (((Ordini.Visualizzato)=True)) ORDER BY Ordini.ID", Connessione, adOpenKeyset, adLockReadOnly
   If rs_ricerca.EOF = False Then
      FineLista = (rs_ricerca.RecordCount >= inMaxOrdini)
   End If
   rs_ricerca.Close
   Set rs_ricerca = Nothing
   Exit Function
errore:
   Set rs_ricerca = Nothing
End Function
Function ProssimoOrdine(ByVal inCodiceinMacc As Integer) As Integer
   Dim rs_ricerca As New ADODB.Recordset
   Dim a
   Dim testTemp As Boolean
   
   On Error GoTo errore
   ProssimoOrdine = 0
   testTemp = False
   rs_ricerca.CursorLocation = adUseClient
   rs_ricerca.Open "SELECT Ordini.ID, Ordini.Visualizzato From Ordini WHERE (((Ordini.Visualizzato)=True)) ORDER BY Ordini.ID", Connessione, adOpenKeyset, adLockReadOnly
   If rs_ricerca.EOF = False Then
      rs_ricerca.MoveLast
      If rs_ricerca("ID") = 99 Then
         a = 0
         Do
            a = rs_ricerca("ID")
            rs_ricerca.MovePrevious
            If rs_ricerca.BOF Then
               testTemp = True
            Else
               a = a - rs_ricerca("ID")
            End If
         Loop Until testTemp = True Or (a > 1)
      End If
      If testTemp = False Then
         ProssimoOrdine = rs_ricerca("ID") + 1
      Else
         ProssimoOrdine = 1
      End If
   End If
   rs_ricerca.Close
   Set rs_ricerca = Nothing
   If ProssimoOrdine = 0 Then ProssimoOrdine = inCodiceinMacc
   Exit Function
errore:
   Set rs_ricerca = Nothing
End Function

Function CopiaDatiinDBRicette(Optional inNuova As Boolean = True) As Boolean
   Dim rs_ricerca As New ADODB.Recordset
   Dim tempOK As Boolean
   Dim strRicerca As String
   Dim i As Integer
   
   On Error GoTo errore
   If inNuova Then
      strRicerca = ""
   Else
      strRicerca = " WHERE ID='" & Trim(DatiFromFile.IDRicetta) & "'"
   End If
   CopiaDatiinDBRicette = False
   rs_ricerca.CursorLocation = adUseClient
   rs_ricerca.Open "SELECT * FROM Ricette" & strRicerca, Connessione, adOpenKeyset, adLockOptimistic
   If rs_ricerca.EOF = False Then
      If inNuova Then
         rs_ricerca.AddNew
         rs_ricerca("ID") = Trim(DatiFromFile.IDRicetta)
      End If
      rs_ricerca("TipoTubo") = DatiFromFile.TipoTubo
      rs_ricerca("Altezza") = DatiFromFile.TAltezza
      rs_ricerca("Larghezza") = DatiFromFile.TLarghezza
      rs_ricerca("Lunghezza") = DatiFromFile.TLunghezza
      rs_ricerca("Spessore") = DatiFromFile.TSpessore
      rs_ricerca("PesoTeoricoTubo") = DatiFromFile.TPeso
      rs_ricerca("TipoPacco") = DatiFromFile.TipoPacco
      rs_ricerca("NumeroTubi") = DatiFromFile.NumeroTubi
      rs_ricerca("NumeroFile") = DatiFromFile.NumeroFile
      rs_ricerca("LarghezzaBasePacco") = DatiFromFile.LargBase
      rs_ricerca("LarghezzaFilaMax") = DatiFromFile.LargMax
      rs_ricerca("LarghezzaLatoPaccoEsagono") = DatiFromFile.LargEsa
      rs_ricerca("AltezzaPacco") = DatiFromFile.AltezzaPacco
      rs_ricerca("FilaUscitaControsagoma") = DatiFromFile.FilaContros
      rs_ricerca("PesoTeoricoPacco") = DatiFromFile.PesoPacco
      rs_ricerca("NumeroRegge") = DatiFromFile.Regge
      rs_ricerca("R_473_64_VelMagneti") = DatiFromFile.VelMagneti
      For i = 1 To 50
         rs_ricerca("Fila" & Format(i, "00")) = DatiFromFile.TubiFila(i)
         If i < 13 Then
            rs_ricerca("Reggia" & Format(i, "00")) = DatiFromFile.Reggia(i)
         End If
      Next
      rs_ricerca.UpdateBatch adAffectCurrent
   End If
   rs_ricerca.Close
   Set rs_ricerca = Nothing
   CopiaDatiinDBRicette = True
   Exit Function
errore:
   On Error Resume Next
   rs_ricerca.CancelBatch adAffectCurrent
   Set rs_ricerca = Nothing
End Function

Function CopiaDatiinDBOrdini(ByVal CodiceOrdine As Integer) As Boolean
   Dim rs_ricerca As New ADODB.Recordset
   Dim tempOK As Boolean
   
   On Error GoTo errore
   CopiaDatiinDBOrdini = False
   rs_ricerca.CursorLocation = adUseClient
   rs_ricerca.Open "SELECT * FROM Ordini WHERE ID=" & Trim(Str(CodiceOrdine)), Connessione, adOpenKeyset, adLockOptimistic
   If rs_ricerca.EOF = False Then
       rs_ricerca("Visualizzato") = True
       rs_ricerca.UpdateBatch adAffectCurrent
       rs_ricerca("IDRicetta") = DatiFromFile.IDRicetta
       rs_ricerca("Descrizione") = DatiFromFile.Descrizione
       rs_ricerca("NumPacchi") = DatiFromFile.NumPacchi
       rs_ricerca("GestioneAFineOrdine") = DatiFromFile.GestFineOrdine
       rs_ricerca.UpdateBatch adAffectCurrent
   End If
   rs_ricerca.Close
   Set rs_ricerca = Nothing
   CopiaDatiinDBOrdini = True
   Exit Function
errore:
   On Error Resume Next
   rs_ricerca.CancelBatch adAffectCurrent
   Set rs_ricerca = Nothing
End Function

Function PopolaUDT(ByVal CodiceOrdine As Integer) As Boolean
   Dim rs_ricerca As New ADODB.Recordset
   Dim tempOK As Boolean
   Dim i As Integer
   Dim strRicerca As String
   
   strRicerca = "SELECT Ordini.ID, Ordini.Descrizione, Ordini.NumPacchi, Ordini.GestioneAFineOrdine, Ricette.* FROM Ordini INNER JOIN Ricette ON Ordini.IDRicetta = Ricette.ID WHERE (Ordini.Id=" & Trim(Str(CodiceOrdine)) & ")"
   On Error GoTo errore
   PopolaUDT = False
   rs_ricerca.CursorLocation = adUseClient
   rs_ricerca.Open strRicerca, Connessione, adOpenKeyset, adLockReadOnly
   If rs_ricerca.EOF = False Then
       DatiFromFile.IDRicetta = rs_ricerca("Ricette.ID")
       DatiFromFile.Descrizione = rs_ricerca("Descrizione")
       DatiFromFile.NumPacchi = rs_ricerca("NumPacchi")
       DatiFromFile.GestFineOrdine = rs_ricerca("GestioneAFineOrdine")
       DatiFromFile.TipoTubo = rs_ricerca("TipoTubo")
       DatiFromFile.TAltezza = rs_ricerca("Altezza")
       DatiFromFile.TLarghezza = rs_ricerca("Larghezza")
       DatiFromFile.TLunghezza = rs_ricerca("Lunghezza")
       DatiFromFile.TSpessore = rs_ricerca("Spessore")
       DatiFromFile.TPeso = rs_ricerca("PesoTeoricoTubo")
       DatiFromFile.TipoPacco = rs_ricerca("TipoPacco")
       DatiFromFile.NumeroTubi = rs_ricerca("NumeroTubi")
       DatiFromFile.NumeroFile = rs_ricerca("NumeroFile")
       DatiFromFile.LargBase = rs_ricerca("LarghezzaBasePacco")
       DatiFromFile.LargMax = rs_ricerca("LarghezzaFilaMax")
       DatiFromFile.LargEsa = rs_ricerca("LarghezzaLatoPaccoEsagono")
       DatiFromFile.AltezzaPacco = rs_ricerca("AltezzaPacco")
       DatiFromFile.FilaContros = rs_ricerca("FilaUscitaControsagoma")
       DatiFromFile.PesoPacco = rs_ricerca("PesoTeoricoPacco")
       DatiFromFile.Regge = rs_ricerca("NumeroRegge")
       DatiFromFile.VelMagneti = rs_ricerca("R_473_64_VelMagneti")
       For i = 1 To 50
          DatiFromFile.TubiFila(i) = rs_ricerca("Fila" & Format(i, "00"))
          If i < 13 Then
             DatiFromFile.Reggia(i) = rs_ricerca("Reggia" & Format(i, "00"))
          End If
       Next
   End If
   rs_ricerca.Close
   Set rs_ricerca = Nothing
   PopolaUDT = True
   Exit Function
errore:
   On Error Resume Next
   Set rs_ricerca = Nothing
End Function

Function ScompattaDati(ByVal inorigine As String) As Boolean
   Dim i As Integer
   
   On Error GoTo errore
   ScompattaDati = False
   With DatiFromFile
      .AltezzaPacco = Left(inorigine, DataLen.lenAltezzapacco)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenAltezzapacco)
      .Descrizione = Left(inorigine, DataLen.lenDescrizione)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenDescrizione)
      .FilaContros = Left(inorigine, DataLen.lenFilaContros)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenFilaContros)
      .GestFineOrdine = Left(inorigine, DataLen.lenGestFineOrdine)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenGestFineOrdine)
      .IDRicetta = Left(inorigine, DataLen.lenIDricetta)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenIDricetta)
      .LargBase = Left(inorigine, DataLen.lenLargBase)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenLargBase)
      .LargEsa = Left(inorigine, DataLen.lenLargEsa)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenLargEsa)
      .LargMax = Left(inorigine, DataLen.lenLargMax)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenLargMax)
      .NumeroFile = Left(inorigine, DataLen.lenNumeroFile)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenNumeroFile)
      .NumeroTubi = Left(inorigine, DataLen.lenNumeroTubi)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenNumeroTubi)
      .NumPacchi = Left(inorigine, DataLen.lenNumPacchi)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenNumPacchi)
      .PesoPacco = Left(inorigine, DataLen.lenPesoPacco)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenPesoPacco)
      .Regge = Left(inorigine, DataLen.lenRegge)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenRegge)
      .TAltezza = Left(inorigine, DataLen.lenTAltezza)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTAltezza)
      .TipoPacco = Left(inorigine, DataLen.lenTipoPacco)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTipoPacco)
      .TipoTubo = Left(inorigine, DataLen.lenTipotubo)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTipotubo)
      .TLarghezza = Left(inorigine, DataLen.lenTLarghezza)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTLarghezza)
      .TLunghezza = Left(inorigine, DataLen.lenTLunghezza)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTLunghezza)
      .TPeso = Left(inorigine, DataLen.lenTPeso)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTPeso)
      .TSpessore = Left(inorigine, DataLen.lenTSpessore)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTSpessore)
      .VelMagneti = Left(inorigine, DataLen.lenVelMagneti)
      inorigine = Right(inorigine, Len(inorigine) - DataLen.lenVelMagneti)
      For i = 1 To 50
         If i < 13 Then
           .Reggia(i) = Left(inorigine, DataLen.lenReggia)
           inorigine = Right(inorigine, Len(inorigine) - DataLen.lenReggia)
         End If
         .TubiFila(i) = Left(inorigine, DataLen.lenTubiFila)
         inorigine = Right(inorigine, Len(inorigine) - DataLen.lenTubiFila)
      Next
   End With
   ScompattaDati = True
errore:
End Function

Function FilesNellaCartella(ByVal inNomeCartella As String, ByVal inTipoFile As String) As Long
   Dim sSource As String
   Dim sFileType As String
   Dim dwAttributes As Long
   Dim numFiles As Long
   
   On Error GoTo errore
   sSource = inNomeCartella 'c:\
   sFileType = inTipoFile '"*.*"
   FilesNellaCartella = 0
   
   dwAttributes = FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_ARCHIVE
   numFiles = FilesCountByAttribute(sSource, sFileType, dwAttributes)
   FilesNellaCartella = numFiles
   MaxCountFiles = numFiles
   
errore:
End Function

Function ElencoFilesNellaCartella(ByVal inNomeCartella As String, ByVal inTipoFile As String) As Long
   Dim sSource As String
   Dim sFileType As String
   Dim dwAttributes As Long
   Dim numFiles As Long
   
   On Error GoTo errore
   sSource = inNomeCartella 'c:\
   sFileType = inTipoFile '"*.*"
   ElencoFilesNellaCartella = 0
   
   dwAttributes = FILE_ATTRIBUTE_ARCHIVE
   numFiles = FilesListByAttribute(sSource, sFileType, dwAttributes)
   ElencoFilesNellaCartella = numFiles
   MaxCountFiles = numFiles
   
errore:
End Function

Private Function FilesCountAll(sSource As String, sFileType As String) As Long

   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long
   Dim fCount As Long
   

   hFile = FindFirstFile(sSource & sFileType, wfd)
   
   If hFile <> INVALID_HANDLE_VALUE Then
      Do
         FilesCountAll = FilesCountAll + 1
      Loop Until FindNextFile(hFile, wfd) = 0
   End If
   Call FindClose(hFile)
   
End Function


Private Function FilesCountByAttribute(sSource As String, _
                                      sFileType As String, _
                                      dwAttributes As Long) As Long

   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long
   
   hFile = FindFirstFile(sSource & sFileType, wfd)
   
   If (hFile <> INVALID_HANDLE_VALUE) Then
      
      Do
         If dwAttributes And GetFileAttributes(sSource & wfd.cFileName) Then
            FilesCountByAttribute = FilesCountByAttribute + 1
         End If
         
      Loop Until FindNextFile(hFile, wfd) = 0
      
   End If
   Call FindClose(hFile)
   
End Function


Private Function FilesListByAttribute(sSource As String, _
                                      sFileType As String, _
                                      dwAttributes As Long) As Long

   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long
   Dim i As Long
   
   i = 0
   ReDim arrayFile(i) As String
   hFile = FindFirstFile(sSource & sFileType, wfd)
   
   If (hFile <> INVALID_HANDLE_VALUE) Then
      Do
         If (dwAttributes And GetFileAttributes(sSource & wfd.cFileName)) = dwAttributes Then
            ReDim Preserve arrayFile(i) As String
            arrayFile(i) = TrimNull(wfd.cFileName)
            FilesListByAttribute = FilesListByAttribute + 1
            i = i + 1
         End If
      Loop Until FindNextFile(hFile, wfd) = 0
   End If
   
   Call FindClose(hFile)
   
End Function


Private Function GetAttributeString(attr As Long) As String

   Dim tmp As String
   
   If attr And FILE_ATTRIBUTE_ARCHIVE Then tmp = tmp & "ARCHIVE  "
   If attr And FILE_ATTRIBUTE_NORMAL Then tmp = tmp & "NORMAL  "
   If attr And FILE_ATTRIBUTE_HIDDEN Then tmp = tmp & "HIDDEN  "
   If attr And FILE_ATTRIBUTE_READONLY Then tmp = tmp & "READONLY  "
   If attr And FILE_ATTRIBUTE_SYSTEM Then tmp = tmp & "SYSTEM  "
   If attr And FILE_ATTRIBUTE_TEMPORARY Then tmp = tmp & "TEMPORARY  "
   If attr And FILE_ATTRIBUTE_COMPRESSED Then tmp = tmp & "COMPRESSED  "
   If attr And FILE_ATTRIBUTE_DIRECTORY Then tmp = tmp & "DIRECTORY  "

   GetAttributeString = tmp

End Function


Private Function TrimNull(startstr As String) As String

   TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
   
End Function

Function ScriveMessaggio(ByVal inFilenome As String, ByVal inMessaggio As String) As Boolean
   Dim a$
   Dim f As Integer
   
   ScriveMessaggio = False
   On Error GoTo errore
   f = FreeFile()
   a$ = inMessaggio
   If a$ = "" Then Exit Function
   Open inFilenome For Output Lock Read Write As #f
   Print #f, a$
   ScriveMessaggio = True
errore:
   On Error Resume Next
   Close #f
End Function

