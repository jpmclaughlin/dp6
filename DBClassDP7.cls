VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Enum TTipoBit
   Blocco = 0
   PlcIN = 1
   PlcOUT = 2
End Enum

Const WRITEASYNC_ID = 1

Public Enum stato
   Attivo = 1
   Disattivo = 0
   Simula = 2
End Enum

Public m_Server As stato
Public NumItems
Public DatiCambiati As Boolean
Public NumCambiamenti As Integer
Public ComError As String
Public LogErrON As Boolean
Public LogDCON As Boolean
Public LetturaEseguita As Boolean ' da azzerare nel client prima del passaggio dati e dopo l'acquisizione

Public ScritturaEseguita As Boolean 'da azzerare nel client prima del passaggio dati e dopo l'acquisizione
Private m_WriteError As Long
Private m_CancelID As Long
Private m_DBid As String
Private m_ItemError As String
Private m_DBerror As String


'Private WithEvents db As OPCGroup
Private db As Object

Private MyItems As New Collection
Private MyItem() As Item_class

Property Get NomeItem(ByVal Index As Variant, Optional ByVal nByte As Integer, Optional ByVal nBit As Byte, Optional ByVal TipoBit As TTipoBit) As Variant
   Dim ItemVal As Item_class
   Dim StrReal As String
   
   On Error Resume Next
   
   '============== fornisce il nome dato l'indice nella collezione ======
   If Index <> "" Then
         If Index >= LBound(MyItem) And Index <= UBound(MyItem) Then
            NomeItem = MyItem(Index).Nome
         End If
   Else
        '=============== Fornisce il nome dato l'indirizzo ========
        On Error GoTo 0
        On Error GoTo ItemNonPresente
        
        Select Case TipoBit
        Case 0
           StrReal = LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
        Case 1
           StrReal = "E" & LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
        Case 2
           StrReal = "A" & LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
        End Select
        
        Set ItemVal = MyItems.Item(StrReal)
        NomeItem = ItemVal.Nome
   End If
   Exit Property
    
ItemNonPresente:

   m_ItemError = m_DBid & ":GET Bit " & Str(nByte) & Str(nBit) & " Not present"
   m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
   If LogErrON Then LogFile LogComPath, m_ItemError
  
End Property
Property Get Server() As stato
  Server = m_Server
End Property
Property Let Server(ByVal Nuovo As stato)
   m_Server = Nuovo
   If m_DBid = "" Then m_DBid = "Not present"
   If m_Server = Attivo Then
      LogFile LogComPath, "<tr class='COM'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_DBid & ":ServerState:" & StatoStr(Nuovo) & "</td></tr>"
   Else
      LogFile LogComPath, "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_DBid & ":ServerState:" & StatoStr(Nuovo) & "</td></tr>"
   End If
   If LogErrON Then LogFile LogComPath, ComError
End Property
Private Function StatoStr(ByVal Nuovo As stato) As String
   Select Case Nuovo
   Case 0
        StatoStr = "OFF"
   Case 1
        StatoStr = "ON"
   Case 2
        StatoStr = "Simul"
   End Select
End Function
' inizializzazione dei gruppi
Sub Init(ByVal IdDB As String, ByVal Collegamento As String, ByVal NomeOgg As String, Optional ByVal ValoreIni As Variant)
   Dim NomeBit As Variant
   
      If Server = Disattivo Then Exit Sub
      If Server = Attivo Then
        If db Is Nothing Then
           On Error Resume Next
           OggServer.OPCGroups.Remove (IdDB)
           Set db = OggServer.OPCGroups.Add(IdDB)
           On Error GoTo 0
'           db.IsSubscribed = True
'           db.IsActive = True
'           db.UpdateRate = 50
           m_DBid = IdDB
        End If
        ' imposta il tempo di update per il gruppi
'        OggServer.OPCGroups.DefaultGroupUpdateRate = 250
      End If
      DoEvents
      ComError = ""
      NumItems = NumItems + 1
      ReDim Preserve MyItem(NumItems)
      Set MyItem(NumItems) = New Item_class
      'nuova collezione
'      On Error GoTo ChiaveDoppia
      MyItems.Add MyItem(NumItems), NomeOgg
      'nuovo item nel server
      If Server = Attivo Then
'         On Error GoTo ErrCollegamento
'         Set MyItem(NumItems).Oggetto = db.OPCItems.AddItem(Collegamento, NumItems)

'==================================================
' aggiungere nuovi dati qui
'==================================================

      '
      Else
         If InStr(NomeOgg, ".") = 0 Then
            MyItem(NumItems).valore = ValoreIni
         Else
            MyItem(NumItems).valore = Abs(ValoreIni)
         End If
      End If
      MyItem(NumItems).Nome = NomeOgg
      Exit Sub

ChiaveDoppia:
        ComError = "Init:Elemento duplicato : DB" & IdDB & "," & NomeOgg
        ComError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & ComError & "</td></tr>"
        If LogErrON Then LogFile LogComPath, ComError
        Exit Sub
ErrCollegamento:
        ComError = "Init:Il collegamento : " & Collegamento & " non è corretto"
        ComError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & ComError & "</td></tr>"
        If LogErrON Then LogFile LogComPath, ComError
End Sub
Property Get Item(ByVal IDnome As Variant) As Variant
   Dim ItemVal As Item_class
    
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
   ' Set ItemVal = MyItems.Item(IDnome)
   ' Item = ItemVal.valore
    
    'MODIFICA DP7
    Item = Val(PLC.Item(IDnome))
    
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
    Exit Property
    
ItemNonPresente:

   m_ItemError = m_DBid & ":GET Item " & IDnome & " Not present"
   m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
   If LogErrON Then LogFile LogComPath, m_ItemError
End Property

Property Let Item(ByVal IDnome As Variant, ByVal NuovoValore As Variant)
     Dim ItemVal As Item_class
     Dim SHandle(1) As Long
     Dim ErrorNr() As Long
     Dim Cancel_Id As Long
     Dim NewValue(1) As Variant
     
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
    m_ItemError = "": ScritturaEseguita = False
    Set ItemVal = MyItems.Item(IDnome)
    If Server = Attivo Then
       SHandle(1) = ItemVal.Oggetto.ServerHandle
       NewValue(1) = NuovoValore
       db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
    Else
       ItemVal.valore = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    End If
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":Item " & IDnome & " " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Property
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET Item " & IDnome & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Property
Property Get OpcByte(ByVal nByte As Integer) As Variant
   Dim ItemVal As Item_class
   Dim StrByte As String
   
   If Server = Disattivo Then Exit Property
   On Error GoTo ItemNonPresente
   StrByte = "b" & LTrim(Str(nByte))
   'Set ItemVal = MyItems.Item(StrByte)
   'OpcByte = Val(ItemVal.valore)
   
    'MODIFICA DP7
    OpcByte = Val(PLC.Item(m_DBid & "," & StrByte))
    
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
   Exit Property
    
ItemNonPresente:

    m_ItemError = m_DBid & ":GET OpcByte " & Str(nByte) & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Property
Property Let OpcByte(ByVal nByte As Integer, ByVal NuovoValore As Variant)
    Dim ItemVal As Item_class
    Dim SHandle(1) As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim NewValue(1) As Variant
    Dim StrIndex As String
     
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
    StrIndex = "b" & LTrim(Str(nByte))
    m_ItemError = "": ScritturaEseguita = False
    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
       SHandle(1) = ItemVal.Oggetto.ServerHandle
       NewValue(1) = NuovoValore
       'db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
    
       PLC.Item(m_DBid & "," & StrIndex) = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    Else
       ItemVal.valore = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    End If
    'errori
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET OpcByte " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Property
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET OpcByte " & Str(nByte) & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Property

Property Get Word(ByVal nword As Integer) As Variant
   Dim ItemVal As Item_class
   Dim strword As String
   
    On Error GoTo ItemNonPresente
    strword = "int" & LTrim(Str(nword))
    
    'Set ItemVal = MyItems.Item(strword)
    'Word = Val(ItemVal.valore)
    
    'MODIFICA DP7
    Word = Val(PLC.Item(m_DBid & "," & strword))
    
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
    Exit Property
    
ItemNonPresente:
   
    m_ItemError = m_DBid & ":GET Word " & Str(nword) & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Property

Property Let Word(ByVal nword As Integer, ByVal NuovoValore As Variant)
     Dim ItemVal As Item_class
     Dim SHandle(1) As Long
     Dim ErrorNr() As Long
     Dim Cancel_Id As Long
     Dim NewValue(1) As Variant
     Dim StrIndex As String
     
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
    StrIndex = "int" & LTrim(Str(nword))
    m_ItemError = "": ScritturaEseguita = False
    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
       'SHandle(1) = ItemVal.Oggetto.ServerHandle
       'NewValue(1) = NuovoValore
       'db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
       
       PLC.Item(m_DBid & "," & StrIndex) = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    Else
       ItemVal.valore = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    End If
    
    If ErrorNr(1) <> 0 Then
       m_ItemError = OggServer.GetErrorString(ErrorNr(1))
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Word " & m_WriteError
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Property
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET Word " & Str(nword) & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
   
End Property

Property Get DWord(ByVal ndword As Integer) As Variant
   Dim ItemVal As Item_class
   Dim StrDWord As String
   
    On Error GoTo ItemNonPresente
    StrDWord = "Dint" & LTrim(Str(ndword))
    'Set ItemVal = MyItems.Item(StrDWord)
    'DWord = Val(ItemVal.valore)
    
    'MODIFICA DP7
    DWord = Val(PLC.Item(m_DBid & "," & StrDWord))
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
    Exit Property
    
ItemNonPresente:

   m_ItemError = m_DBid & ":GET Dword " & Str(ndword) & " Not present"
   m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Property


Property Let DWord(ByVal ndword As Integer, ByVal NuovoValore As Variant)
    Dim ItemVal As Item_class
    Dim SHandle(1) As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim NewValue(1) As Variant
    Dim StrIndex As String
     
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
    StrIndex = "dint" & LTrim(Str(ndword))
    m_ItemError = "": ScritturaEseguita = False
    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
       'SHandle(1) = ItemVal.Oggetto.ServerHandle
       'NewValue(1) = NuovoValore
       'db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
       
       PLC.Item(m_DBid & "," & StrIndex) = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    Else
       ItemVal.valore = NuovoValore
       If NumCambiamenti < 32000 Then
           NumCambiamenti = NumCambiamenti + 1
        Else
           NumCambiamenti = 1
        End If
        DatiCambiati = True
    End If
    If ErrorNr(1) <> 0 Then
       m_ItemError = OggServer.GetErrorString(ErrorNr(1))
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET DWord " & m_WriteError
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Property
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET DWord " & ndword & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError

End Property

Property Get Real(ByVal nreal As Integer) As Variant
   Dim ItemVal As Item_class
   Dim StrReal As String
   
    On Error GoTo ItemNonPresente
    StrReal = "real" & LTrim(Str(nreal))
    'Set ItemVal = MyItems.Item(StrReal)
    'Real = Val(ItemVal.valore)
    
    'MODIFICA DP7
    Real = Val(PLC.Item(m_DBid & "," & StrReal))
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
    Exit Property
    
ItemNonPresente:

   m_ItemError = m_DBid & ":GET Real " & Str(nreal) & " Not present"
   m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError

End Property
Property Let Real(ByVal nreal As Integer, ByVal NuovoValore As Variant)
    Dim ItemVal As Item_class
    Dim SHandle(1) As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim NewValue(1) As Variant
    Dim StrIndex As String
     
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
    StrIndex = "real" & LTrim(Str(nreal))
    m_ItemError = "": ScritturaEseguita = False
    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
       'SHandle(1) = ItemVal.Oggetto.ServerHandle
       'NewValue(1) = NuovoValore
       'db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
       PLC.Item(m_DBid & "," & StrIndex) = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    Else
       ItemVal.valore = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    End If
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Real " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Property
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET Real " & Str(nreal) & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
   
End Property
Property Get Bit(ByVal nByte As Integer, ByVal nBit As Byte, Optional ByVal TipoBit As TTipoBit) As Variant
   Dim ItemVal As Item_class
   Dim StrReal As String
   
   On Error GoTo ItemNonPresente
   
   Select Case TipoBit
   Case 0
      StrReal = LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
   Case 1
      StrReal = "E" & LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
   Case 2
      StrReal = "A" & LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
   End Select
   
   'Set ItemVal = MyItems.Item(StrReal)
   'Bit = ItemVal.valore
   
   'MODIFICA DP7
   Bit = Val(PLC.Item(m_DBid & "," & StrReal))
   
'   If Server = Attivo Then Exit Property
   If Bit <> 0 Then
      Bit = True
   Else
      Bit = False
   End If
   
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
   Exit Property
    
ItemNonPresente:

   m_ItemError = m_DBid & ":GET Bit " & Str(nByte) & Str(nBit) & " Not present"
   m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
   If LogErrON Then LogFile LogComPath, m_ItemError

End Property
Property Let Bit(ByVal nByte As Integer, ByVal nBit As Byte, Optional ByVal TipoBit As TTipoBit, ByVal NuovoValore As Variant)
    Dim ItemVal As Item_class
    Dim SHandle(1) As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim NewValue(1) As Variant
    Dim StrIndex As String
     
    If Server = Disattivo Then Exit Property
    On Error GoTo ItemNonPresente
    m_ItemError = "": ScritturaEseguita = False
    
    Select Case TipoBit
    Case 0
      StrIndex = LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
    Case 1
      StrIndex = "E" & LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
    Case 2
      StrIndex = "A" & LTrim(Str(nByte)) & "." & LTrim(Str(nBit))
    End Select

    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
       'SHandle(1) = ItemVal.Oggetto.ServerHandle
       'NewValue(1) = NuovoValore
       'db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
       PLC.Item(m_DBid & "," & StrIndex) = Abs(NuovoValore)
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    Else
       ItemVal.valore = NuovoValore
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
    End If
    
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Bit " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Property
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET Bit " & Str(nByte) & Str(nBit) & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Property


Property Get MaskBit(ByVal nword As Integer, ByVal nBit As Byte, Optional ByVal ExistTest As Boolean) As Variant
   Dim ItemVal As Item_class
   Dim StrInt As String
   Dim valore As Variant
   
    If nBit >= 0 And nBit < 16 Then
        On Error GoTo ItemNonPresente
        StrInt = "int" & LTrim(Str(nword))
        'Set ItemVal = MyItems.Item(StrInt)
'        If nBit < 8 Then
'           MaskBit = ItemVal.valore And 2 ^ (8 + nBit)
'        Else
'           MaskBit = ItemVal.valore And 2 ^ (nBit - 8)
'        End If
        
        'MODIFICA DP7
        valore = Val(PLC.Item(m_DBid & "," & StrInt))
        If nBit < 8 Then
           MaskBit = valore And 2 ^ (8 + nBit)
        Else
           MaskBit = valore And 2 ^ (nBit - 8)
        End If

        If MaskBit > 0 Then
           MaskBit = True
        Else
           MaskBit = False
        End If
    End If
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
    Exit Property
    
ItemNonPresente:
        
 If ExistTest = False Then
      m_ItemError = m_DBid & ":GET Mask bit " & Str(nword) & Str(nBit) & " Not present"
      m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
      If LogErrON Then LogFile LogComPath, m_ItemError
   Else
      MaskBit = 2
   End If
   
End Property

Property Let MaskBit(ByVal nword As Integer, ByVal nBit As Byte, Optional ByVal ExistTest As Boolean, ByVal NuovoValore As Variant)
        Dim ItemVal As Item_class
        Dim SHandle(1) As Long
        Dim ErrorNr() As Long
        Dim Cancel_Id As Long
        Dim NewValue(1) As Variant
        Dim StrIndex As String
        Dim ValWord As Variant
        Dim StatoBit As Variant
         
        If Server = Disattivo Then Exit Property
        On Error Resume Next
        If nBit < 0 Or nBit > 16 Then Exit Property
        m_ItemError = "": ScritturaEseguita = False
        StrIndex = "int" & LTrim(Str(nword))
        Set ItemVal = MyItems.Item(StrIndex)
        If ItemVal Is Nothing Then StrIndex = "w" & LTrim(Str(nword))
        Set ItemVal = MyItems.Item(StrIndex)
        On Error GoTo 0
        If ItemVal Is Nothing Then GoTo ItemNonPresente
        
        ' setta o resetta il bit
        If CBool(NuovoValore) Then
            If nBit < 8 Then
                NuovoValore = ValWord Or 2 ^ (8 + nBit)
            Else
                NuovoValore = ValWord Or 2 ^ (nBit - 8)
            End If
            StatoBit = 1
        Else
            If nBit < 8 Then
                 StatoBit = ValWord And 2 ^ (8 + nBit)
            Else
               StatoBit = ValWord And 2 ^ (nBit - 8)
            End If
            NuovoValore = ValWord - StatoBit
        End If
        ' scrive il bit
        If StatoBit > 0 Then
           If Server = Attivo Then
                PLC.Item(m_DBid & "," & StrIndex) = NuovoValore
                If NumCambiamenti < 32000 Then
                   NumCambiamenti = NumCambiamenti + 1
                Else
                   NumCambiamenti = 1
                End If
                DatiCambiati = True
       
'              SHandle(1) = ItemVal.Oggetto.ServerHandle
'              NewValue(1) = NuovoValore
'              db.AsyncWrite 1, SHandle, NewValue, ErrorNr, WRITEASYNC_ID, Cancel_Id
           Else
              ItemVal.valore = NuovoValore
              If NumCambiamenti < 32000 Then
                  NumCambiamenti = NumCambiamenti + 1
              Else
                  NumCambiamenti = 1
              End If
              DatiCambiati = True
           End If
           If ErrorNr(1) <> 0 Then
               m_WriteError = ErrorNr(1)
               m_ItemError = m_DBid & ":LET MaskBit " & OggServer.GetErrorString(ErrorNr(1))
               m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
               If LogErrON Then LogFile LogComPath, m_ItemError
           End If
           m_CancelID = Cancel_Id
           Erase ErrorNr
        End If
        Exit Property
       
ItemNonPresente:

        m_ItemError = m_DBid & ":LET MaskBit " & Str(nword) & Str(nBit) & " Not present"
        m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
        If LogErrON Then LogFile LogComPath, m_ItemError
   
End Property

Sub LetturaAsync(ByVal strItem As String)
    Dim ItemVal As Item_class
    Dim SHandle(1) As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim StrIndex As String
     
    If Server = Disattivo Then Exit Sub
    On Error GoTo ItemNonPresente
    StrIndex = strItem
    m_ItemError = ""
    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
'       PLC.Item(m_DBid & "," & StrIndex) = NewValue
       If NumCambiamenti < 32000 Then
          NumCambiamenti = NumCambiamenti + 1
       Else
          NumCambiamenti = 1
       End If
       DatiCambiati = True
'       SHandle(1) = ItemVal.Oggetto.ServerHandle
'       db.AsyncRead 1, SHandle, ErrorNr, READASYNC_ID, Cancel_Id
    Else
       LetturaEseguita = True
    End If

    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Bit " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    LetturaEseguita = False
    Erase ErrorNr
    Exit Sub
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":AsyncRead " & strItem & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
      
End Sub
Sub BlockReadAsync()
    Dim ItemVal As Item_class
    Dim SHandle() As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim StrIndex As String
    Dim i As Integer
    
    On Error GoTo ItemNonPresente
    
    If Server = Disattivo Then Exit Sub
    m_ItemError = ""
    If Server = Attivo Then
       For i = 1 To NumItems
          Set ItemVal = MyItems.Item(i)
          ReDim Preserve SHandle(i)
          SHandle(i) = ItemVal.Oggetto.ServerHandle
       Next
       db.AsyncRead NumItems, SHandle, ErrorNr, READASYNC_ID, Cancel_Id
    Else
       LetturaEseguita = True
    End If
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Bit " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    LetturaEseguita = False
    Erase ErrorNr
    Exit Sub
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":BlockAsyncRead " & i & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
      
End Sub
Sub Refresh()
  Dim Cancel_Id As Long
    
  If Server = Disattivo Then Exit Sub
  If Server = Attivo Then
     On Error Resume Next
'     db.AsyncRefresh OPC_DS_CACHE, REFRESHASYNC_ID, Cancel_Id
    If NumCambiamenti < 32000 Then
       NumCambiamenti = NumCambiamenti + 1
    Else
       NumCambiamenti = 1
    End If
    DatiCambiati = True
  Else
     If NumCambiamenti < 32000 Then
        NumCambiamenti = NumCambiamenti + 1
     Else
        NumCambiamenti = 1
     End If
     DatiCambiati = True
  End If
End Sub

Property Get DB_ID() As String
    DB_ID = m_DBid
End Property
Property Let DB_ID(ByVal Nuovo As String)
    m_DBid = Nuovo
End Property
Private Sub Class_Initialize()
    NumItems = 0
    NumCambiamenti = 0
End Sub

Private Sub Class_Terminate()
    If Server = Disattivo Then Exit Sub
      If Server = Attivo Then
           On Error Resume Next
           OggServer.OPCGroups.Remove (m_DBid)
     End If
End Sub

Private Sub db_AsyncReadComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date, Errors() As Long)
         Dim i As Integer
                                                                                                                
        On Error Resume Next
        m_DBerror = "GOOD"
        For i = 1 To NumItems
           MyItem(ClientHandles(i)).valore = ItemValues(i)
           If Qualities(i) <> OPC_QUALITY_GOOD Then
              m_DBerror = "BAD (" & MyItem(ClientHandles(i)).Nome & ")"
           End If
        Next
        LetturaEseguita = True
End Sub

Private Sub DB_AsyncWriteComplete(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, Errors() As Long)
   ScritturaEseguita = True
End Sub

'Private Sub DB_DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
'        Dim i As Integer
'
'        m_DBerror = "GOOD"
'        For i = 1 To NumItems
'           MyItem(ClientHandles(i)).valore = ItemValues(i)
'           If Qualities(i) <> OPC_QUALITY_GOOD Then m_DBerror = "BAD (" & MyItem(ClientHandles(i)).Nome & ")"
'           If LogDCON Then
'              LogFile LogComPath, "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_DBid & ":DataChange " & MyItem(ClientHandles(i)).Nome & "=" & ItemValues(i) & "</td></tr>"
'           End If
'        Next
'        If NumCambiamenti < 32000 Then
'           NumCambiamenti = NumCambiamenti + 1
'        Else
'           NumCambiamenti = 1
'        End If
'        DatiCambiati = True
'End Sub


'Property Get NomeItem(ByVal IDnome As Integer) As String
'    Dim ItemVal As Item_class
'
'    On Error GoTo ItemNonPresente
'    Set ItemVal = MyItems.Item(IDnome)
'    NomeItem = ItemVal.Nome
'    Exit Property
'
'ItemNonPresente:
'
'   m_ItemError = m_DBid & ":GET NomeItem " & IDnome & " Not present"
'   m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
'    If LogErrON Then LogFile LogComPath, m_ItemError
'
'End Property

Property Get ErroreItem() As String
   ErroreItem = m_ItemError
End Property

' Ci sono degli ITEMS non validi o non è più attiva la comunicazione
'
Property Get ErroreDB() As String
    ErroreDB = m_DBerror
    'If m_DBerror = "" Then m_DBerror = "Connect OFF"
    'If LogErrON And ComError = "" Then
    '    If m_DBerror = "GOOD" Then
    '      LogFile LogComPath, "<tr class='COM'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_DBid & ":Status DB(I/O Items):" & m_DBerror & "</td></tr>"
    '    Else
    '      LogFile LogComPath, "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_DBid & ":Status DB(I/O Items):" & m_DBerror & "</td></tr>"
    '    End If
   'End If
End Property

Private Function funcErrore(ByVal Qualità As Long) As String
    Select Case Qualità
        Case 0
             funcErrore = "Bad"
        Case 64
             funcErrore = "Uncertain"
        Case 192
             funcErrore = "GOOD"
        Case 8
             funcErrore = "Not_connected"
        Case 13
             funcErrore = "Device_failure"
        Case 16
             funcErrore = "sensor_failure"
        Case 20
             funcErrore = "last_known"
        Case 24
             funcErrore = "comm_failure"
        Case 28
             funcErrore = "out of service"
        Case 132
             funcErrore = "last usable"
        Case 144
             funcErrore = "sensor cal"
        Case 148
             funcErrore = "egu exceeded"
        Case 152
             funcErrore = "sub normal"
        Case 216
             funcErrore = "local override"
        Case Else
             funcErrore = "unknown error"
        End Select
End Function

Sub CancLogFile(ByVal path As String)
   On Error Resume Next
   Kill path
End Sub

Sub LogFile(ByVal path As String, ByVal Err As String, Optional Cancel As Boolean)
Static NFileErr
'Static Errs(0 To 200)
Static Nerr As Integer
Dim VerGood As Integer

    On Error Resume Next
     
     NFileErr = NFileErr + 1
     If NFileErr > 2500 Or Cancel Then
        NFileErr = 0
        Kill path
         On Error GoTo 0
        On Error GoTo Errore
        DoEvents
        Open path For Output As #1
        Print #1, "<html><head><title> DP6 LOG </title><meta name='generator' content='Log generated' /><meta name='author' content='Leonardi Alberto' />"
        Print #1, "<style>"
        Print #1, ".COM{COLOR:green;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt} .xCOM{COLOR:green;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}"
        Print #1, ".DBerr{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt} .xDBerr{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}"
        Print #1, "</style><script id='clientEventHandlersJS' language='javascript'>"
        Print #1, "function ChangeButton(MyButton, MyStyle, ValueOn, ValueOff){if(document.styleSheets[0].rules[MyStyle].style.display=='none'){document.styleSheets[0].rules[MyStyle].style.display='';document.all[MyButton].value=ValueOn;}else{document.styleSheets[0].rules[MyStyle].style.display='none';document.all[MyButton].value=ValueOff;}}"
        
        Print #1, "function btn0_onclick(){ChangeButton('btn0', 0, 'Hide COM good', 'Show COM good');}"
        Print #1, "function btn1_onclick(){ChangeButton('btn1', 2, 'Hide DB err', 'Show DB err');}"
        
        Print #1, "</script></head><body>"
        Print #1, "<h1 style='font-family:Arial' style='font-size:16pt'>" & Date$ & "   :  DP6.0 - ErrorLog</h1><P>"
        Print #1, "<font size=2>"
        Print #1, "<table style='font-family:Arial' style='font-size:10pt' bgcolor='white' align='justify' cellspacing='1' width='100%' border='0' borderColorDark='white' borderColorLight='white' background='' borderColor='white'>"
        Print #1, "<tr><td bgcolor='#e0e0e0' width='100'>EXE - File</td><td width='*'>" & App.EXEName & "</td></tr> "
        Print #1, "<tr><td bgcolor='#e0e0e0' width='100'>EXE - Path</td><td width='*'>" & App.path & "</td></tr> "
        Print #1, "<tr><td bgcolor='#e0e0e0' width='100'>EXE - Ver </td><td width='*'>" & App.Major & "." & App.Minor & "." & App.Revision & "</td></tr> "
        Print #1, "<tr><td bgcolor='#e0e0e0' width='100'>EXE - Trade </td><td width='*'>" & App.LegalCopyright & "</td></tr> "
        Print #1, "<tr><td bgcolor='#e0e0e0' width='100'>EXE - ThID </td><td width='*'>" & App.ThreadID & "</td></tr> "
        Print #1, "</table><hr /><table style='font-family:Arial' style='FONT-SIZE:10pt' border='0' align='justify' cellspacing='1' width='576'><tr>"
        
        Print #1, "<td width='96'><BUTTON id='btn0' class='xCOM' onclick='btn0_onclick()' style='WIDTH: 105px'>Hide COM good</BUTTON></td></tr></table>"
        Print #1, "<td width='96'><BUTTON id='btn1' class='xDBerr' onclick='btn1_onclick()' style='WIDTH: 105px'>Hide DB err</BUTTON></td></tr></table>"
        Print #1, "<hr />"
        Print #1, "<table cols='2' style='font-family:Arial' style='font-size:10pt' border='0' align='justify' cellspacing='1' width='100%'>"
     Else
'        Errs(Nerr) = Err
'        Nerr = Nerr + 1
'        If Nerr > UBound(Errs) Then
'           Nerr = 0
           Open path For Append As #1
'           For Nerr = LBound(Errs) To UBound(Errs)
              'Print #1, Errs(Nerr)
              Print #1, Err
'           Next
'        End If
    End If
    Close #1
    Exit Sub

Errore:
   DoEvents
   Cancel = True
   Close #1
End Sub

'========================================================================
Function LetturaSync(ByVal strItem As String) As Variant
     Dim ItemVal As Item_class
     Dim SHandle(1) As Long
     Dim ErrorNr() As Long
     Dim Cancel_Id As Long
     Dim StrIndex As String
     Dim Valori() As Variant
    
    If Server = Disattivo Then Exit Function
    On Error GoTo ItemNonPresente
    StrIndex = strItem
    m_ItemError = ""
    Set ItemVal = MyItems.Item(StrIndex)
    If Server = Attivo Then
       SHandle(1) = ItemVal.Oggetto.ServerHandle
       db.SyncRead OPC_DS_DEVICE, 1, SHandle, Valori, ErrorNr
       LetturaSync = Valori(1)
    Else
       LetturaSync = ItemVal.valore
    End If
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Bit " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    LetturaEseguita = False
    Erase ErrorNr
    Exit Function
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":AsyncRead " & strItem & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
      
End Function
Function ScritturaSync(ByVal IDnome As Variant, ByVal NuovoValore As Variant, Optional ByVal Lettura As Boolean) As Variant
    Dim ItemVal As Item_class
    Dim SHandle(1) As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim NewValue(1) As Variant
     
    If Server = Disattivo Then Exit Function
    On Error GoTo ItemNonPresente
    m_ItemError = "": ScritturaEseguita = False
    Set ItemVal = MyItems.Item(IDnome)
    If Server = Attivo Then
       SHandle(1) = ItemVal.Oggetto.ServerHandle
       NewValue(1) = NuovoValore
       db.SyncWrite 1, SHandle, NewValue, ErrorNr
    Else
       ScritturaEseguita = True
    End If
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":Item " & IDnome & " " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    'legge il dato scritto se abilitato
    If Lettura Then ScritturaSync = LetturaSync(IDnome)
    m_CancelID = Cancel_Id
    Erase ErrorNr
    Exit Function
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":LET Item " & IDnome & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
End Function

Sub WORDSReadAsync(ByVal Prima As Integer, ByVal NumItems As Integer)
    Dim ItemVal As Item_class
    Dim SHandle() As Long
    Dim ErrorNr() As Long
    Dim Cancel_Id As Long
    Dim StrIndex As String
    Dim i As Integer
    Dim strword As String
    Dim j As Integer
   
    If Server = Disattivo Then Exit Sub
    On Error GoTo ItemNonPresente
    j = 1
    m_ItemError = ""
    If Server = Attivo Then
       For i = 0 To ((NumItems - 1) * 2) Step 2
          strword = "int" & LTrim(Str(Prima + i))
          Set ItemVal = MyItems.Item(strword)
          ReDim Preserve SHandle(j)
          SHandle(j) = ItemVal.Oggetto.ServerHandle
          j = j + 1
       Next
       db.AsyncRead NumItems, SHandle, ErrorNr, READASYNC_ID, Cancel_Id
    Else
       LetturaEseguita = True
    End If
    If ErrorNr(1) <> 0 Then
       m_WriteError = ErrorNr(1)
       m_ItemError = m_DBid & ":LET Bit " & OggServer.GetErrorString(ErrorNr(1))
       m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
       If LogErrON Then LogFile LogComPath, m_ItemError
    End If
    m_CancelID = Cancel_Id
    LetturaEseguita = False
    Erase ErrorNr
    Exit Sub
        
ItemNonPresente:
    
    m_ItemError = m_DBid & ":BlockAsyncRead " & i & " Not present"
    m_ItemError = "<tr class='DBerr'><td width='85'>" & Time$ & "</td><td width='*'><B>" & m_ItemError & "</td></tr>"
    If LogErrON Then LogFile LogComPath, m_ItemError
      
End Sub
' lettura multipla di dati all'interno del gruppo
Property Let MultiReadValore(ByVal Index As Long, ByVal valore As Variant)
    If Index <= NumItems Then
        MyItem(Index).valore = valore
    End If
End Property

Property Get MultiReadValore(ByVal Index As Long)
    If Index <= NumItems Then MultiReadValore = MyItem(Index).valore
End Property


