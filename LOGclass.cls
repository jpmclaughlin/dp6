VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LOGclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private priv_arrAll() As String

Public Codice As String
Public Turno As String
Public Data As String
Public Ora As String
Public stato As String

Function Scrivi(inFilenome As String) As Boolean
  Dim f As Integer, strTemp As String
  
  Scrivi = False
  Err.Clear
  f = FreeFile()
  On Error GoTo Errore
  strTemp = Codice & "|" & Data & "|" & Ora & "|" & Turno & "|" & stato & "|" '& vbCrLf
  If IIf(Dir(inFilenome, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "", True, False) Then
     Open inFilenome For Append As #f
  Else
     Open inFilenome For Output As #f
  End If
  Print #f, strTemp
  
Errore:
  Close #f
  Scrivi = True
  If Err <> 0 Then
     Scrivi = GestioneErrore(inFilenome)
  End If
End Function
Function ContaElementiFile(inFilenome As String) As Long
  Dim f As Integer, strTemp As String
  Dim arrStr() As String
  
  ContaElementiFile = 0
  Err.Clear
  f = FreeFile()
  On Error GoTo Errore
  If IIf(Dir(inFilenome, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "", True, False) = False Then Exit Function
     Open inFilenome For Input As #f
     strTemp = Input(LOF(f), #f)
Errore:
  Close #f
  If Err <> 0 Then
     ContaElementiFile = Abs(GestioneErrore(inFilenome))
  End If
  ' conta elementi
  arrStr = Split(strTemp, vbCrLf)
  ContaElementiFile = UBound(arrStr)
End Function
Function removeElementiFile(inFilenome As String, ByVal inMaxelem As Long, Optional ByVal inDaCanc As Long) As Boolean
  Dim f As Integer, strTemp As String, elementiDacancellare As Long, ContaElementi As Long
  Dim i As Integer
  Dim arrStr() As String
  
  On Error GoTo Errore
  removeElementiFile = False
  Err.Clear
  f = FreeFile()
  If IIf(Dir(inFilenome, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "", True, False) = False Then Exit Function
  Open inFilenome For Input As #f
  strTemp = Input(LOF(f), #f)
  Close #f
  arrStr = Split(strTemp, vbCrLf)
  ContaElementi = UBound(arrStr)
  If IsMissing(inDaCanc) Or inDaCanc = 0 Then inDaCanc = 100
  If inMaxelem <= inDaCanc Then inMaxelem = inDaCanc + 10
  If ContaElementi <= inMaxelem Then Exit Function
  strTemp = ""
  For i = ContaElementi - inDaCanc To ContaElementi
     strTemp = strTemp & arrStr(i) & vbCrLf
  Next
  Open inFilenome For Output As #f
  Print #f, strTemp
  removeElementiFile = True
Errore:
  Close #f
  If Err <> 0 Then
     removeElementiFile = GestioneErrore(inFilenome)
  End If
End Function

Function Rspopolato(ByVal inNomefile As String, ParamArray inCheckLen() As Variant) As ADODB.Recordset
  Dim rsTemp As New ADODB.Recordset
  Dim dataCount As Long, dataTotCount As Long
  Dim tempArr() As String
  Dim i As Long
  Dim j As Integer
  Dim vtest
  
  On Error GoTo Errore
    
  'determina lunghezza della riga
  dataCount = UBound(inCheckLen)
  dataTotCount = 0
  For i = 0 To dataCount
     dataTotCount = dataTotCount + inCheckLen(i)
  Next
  
  If Leggi(inNomefile) = False Then Exit Function
     
  With rsTemp
    If .State = adStateOpen Then .Close
    .Fields.Append "Data", adChar, 12
    .Fields.Append "Ora", adChar, 12
    .Fields.Append "Turno", adChar, 100
    .Fields.Append "Stato", adChar, 5
    .Fields.Append "Descrizione", adChar, 12
    .Open
    i = 0
    Do
        priv_arrAll(i) = Replace(priv_arrAll(i), " ", "")
        priv_arrAll(i) = Replace(priv_arrAll(i), vbCr, "")
        priv_arrAll(i) = Replace(priv_arrAll(i), vbCrLf, "")
        vtest = Replace(priv_arrAll(i), "|", "")
        If Len(vtest) <> dataTotCount Then
           GoTo FineLettura
        Else
           tempArr = Split(priv_arrAll(i), "|")
        End If
        .AddNew
        j = 0
        If Len(tempArr(j)) = inCheckLen(j) Then .Fields("Descrizione") = Trim(tempArr(j))
        j = j + 1
        If Len(tempArr(j)) = inCheckLen(j) Then .Fields("Data") = Trim(tempArr(j))
        j = j + 1
        If Len(tempArr(j)) = inCheckLen(j) Then .Fields("Ora") = Trim(tempArr(j))
        j = j + 1  ' nella rappresentazione nella griglia convertire in lettere
        If Len(tempArr(j)) = inCheckLen(j) Then .Fields("Turno") = Trim(tempArr(j))
        j = j + 1
        If Len(tempArr(j)) = inCheckLen(j) Then .Fields("Stato") = Trim(tempArr(j))
        .MoveNext
        i = i + 1
    Loop Until i >= (UBound(priv_arrAll) + 1)
  End With
FineLettura:
  Set Rspopolato = rsTemp
 ' rsTemp.Close
 ' Set rsTemp = Nothing
Errore:

End Function

Function GestioneErrore(inFilenome As String) As Boolean
   Dim f As Integer
   
   Err.Clear
   f = FreeFile()
   On Error GoTo Errore
   GestioneErrore = False
   SetAttr inFilenome, vbArchive
   Kill inFilenome
   GestioneErrore = True
Errore:
   Err.Clear
   On Error Resume Next
   Open inFilenome For Output As #f
   Close #f
   If Err <> 0 Then GestioneErrore = True
End Function
Function Leggi(inFilenome As String) As Boolean
  Dim f As Integer, strTemp As String
  Dim arrTemp() As String
  Dim i As Long
  
  Leggi = False
  Erase priv_arrAll
  Err.Clear
  f = FreeFile()
  On Error GoTo Errore
  If IIf(Dir(inFilenome, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "", True, False) = False Then Exit Function
  Open inFilenome For Input As #f
  strTemp = Input(LOF(f), #f)
  Close #f
  priv_arrAll = Split(strTemp, vbCrLf)
  For i = 0 To UBound(priv_arrAll)
      If Len(priv_arrAll(i)) = 0 Then priv_arrAll(i) = vbNullChar
  Next
  priv_arrAll() = Filter(priv_arrAll(), vbNullChar, False)
  Leggi = True
Errore:
  Close #f
  If Err <> 0 Then
     GestioneErrore (inFilenome)
  End If
End Function

Function TestoExport(inValori() As String, inNomefile As String, Optional delim As String = vbCrLf, Optional inapp As Boolean = True) As Boolean
   Dim i As Long
   Dim a
   
   TestoExport = False
   On Error GoTo Errore
   ScriveTextFile Join(inValori, delim), inNomefile, inapp
   TestoExport = True
Errore:
   a = Err.Description
End Function
Function TestoImport(inFilenome As String, Optional delim As String = "|") As Variant()
   Dim lines() As String, i As Long
    
   On Error GoTo Errore
   lines() = Split(LeggeTextFile(inFilenome), vbCrLf) 'car. rit carr.+linefeed
   For i = 0 To UBound(lines)
      If Len(lines(i)) = 0 Then lines(i) = vbNullChar
   Next
   lines() = Filter(lines(), vbNullChar, False)
   ReDim Valori(0 To UBound(lines)) As Variant
   For i = 0 To UBound(lines)
      Valori(i) = Split(lines(i), delim)
   Next
   TestoImport = Valori()
   Exit Function
Errore:
   End Function

Function addNewErr(intesto As String, instato As Boolean, _
                        ByVal t1in As Integer, ByVal t1fn As Integer, _
                        ByVal t2in As Integer, ByVal t2fn As Integer, _
                        ByVal t3in As Integer, ByVal t3fn As Integer) As Boolean
    Dim Oratemp As Integer
    Dim Trovato As Boolean
    Dim i As Integer
    Dim turni(3, 2) As Integer
    
    turni(1, 1) = t1in
    turni(1, 2) = t1fn
    turni(2, 1) = t2in
    turni(2, 2) = t2fn
    turni(3, 1) = t3in
    turni(3, 2) = t3fn
    On Error GoTo Errore
    Oratemp = Format(Hour(Time), "00")
    addNewErr = False
    Codice = Trim(intesto)
    stato = Trim(Str(Abs(instato)))
    Data = Trim(Format(Month(Now), "00") & "/" & Format(Day(Now), "00") & "/" & year(Now))
    Ora = Trim(Format(Oratemp, "00") & ":" & Format(Minute(Time), "00"))
    ' scrive il turno
    Trovato = False
    For i = 1 To 3
         If Hour(Time) >= turni(i, 1) And Hour(Time) < turni(i, 2) Then
            Turno = Trim(Str(i))
            Trovato = True
            Exit For
         End If
     Next
     If Trovato = False Then
         For i = 1 To 3
            If (Hour(Time) > turni(i, 1) And Hour(Time) <= 23) Or (Hour(Time) >= 0 And Hour(Time) < turni(i, 2)) Then
               Turno = Trim(Str(i))
               Exit For
            End If
         Next
     End If
     addNewErr = True
Errore:

End Function
Function TestFile(ByVal inNomefile As String, ParamArray inCheckLen() As Variant) As Boolean
  Dim rsTemp As New ADODB.Recordset
  Dim dataCount As Long, dataTotCount As Long
  Dim tempArr() As String
  Dim tempCheckArr() As String
  Dim i As Long
  Dim j As Integer
  Dim vtest
  
  On Error GoTo Errore
  TestFile = False
  'determina lunghezza della riga
  dataCount = UBound(inCheckLen)
  dataTotCount = 0
  For i = 0 To dataCount
     dataTotCount = dataTotCount + inCheckLen(i)
  Next
  ReDim tempArr(0) As String
  If Leggi(inNomefile) = False Then Exit Function
  i = 0
  Do
     priv_arrAll(i) = Replace(priv_arrAll(i), " ", "")
     priv_arrAll(i) = Replace(priv_arrAll(i), vbCrLf, "")
     priv_arrAll(i) = Replace(priv_arrAll(i), vbCr, "")
     priv_arrAll(i) = Replace(priv_arrAll(i), vbLf, "")
     vtest = Replace(priv_arrAll(i), "|", "")
     If Len(vtest) <> dataTotCount Then
        GoTo FineLettura
     Else
        tempCheckArr = Split(priv_arrAll(i), "|")
        j = 0
        If Len(tempCheckArr(j)) <> inCheckLen(j) Then GoTo FineLettura
        j = j + 1
        If Len(tempCheckArr(j)) <> inCheckLen(j) Then GoTo FineLettura
        j = j + 1
        If Len(tempCheckArr(j)) <> inCheckLen(j) Then GoTo FineLettura
        j = j + 1
        If Len(tempCheckArr(j)) <> inCheckLen(j) Then GoTo FineLettura
        j = j + 1
        If Len(tempCheckArr(j)) <> inCheckLen(j) Then GoTo FineLettura
        ReDim Preserve tempArr(i) As String
        tempArr(i) = priv_arrAll(i)
     End If
     i = i + 1
  Loop Until i >= (UBound(priv_arrAll) + 1)
  TestFile = True
  Exit Function
FineLettura:
  TestFile = TestoExport(tempArr, inNomefile, , False)
  
Errore:
'           tempArr = Split(priv_arrAll(i), "|")
'        End If
'        .AddNew

'        .MoveNext
'        i = i + 1
'    Loop Until i >= (UBound(priv_arrAll) + 1)
End Function
Function LeggeTextFile(inNomefile As String) As String
    Dim fnum As Integer, isopen As Boolean
    
    On Error GoTo Errore
    fnum = FreeFile()
    Open inNomefile For Input As #fnum
    isopen = True
    LeggeTextFile = Input(LOF(fnum), fnum)
    
Errore:
    If isopen Then Close #fnum
    If Err <> 0 Then
       If GestioneErrore(inNomefile) = False Then
          MsgBox "Allarm file corrupt (delete it from target folder)", vbCritical, "DP6"
       End If
    End If
End Function

Function ScriveTextFile(intesto As String, infilename As String, Optional inAppendmode As Boolean)
    Dim Appendmode As Boolean
    Dim fnum As Integer, isopen As Boolean
    
    On Error GoTo Errore
    Appendmode = False
    If IsMissing(inAppendmode) Then
        If IIf(Dir(infilename, vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive) <> "", True, False) Then
           Appendmode = True
        End If
    Else
        Appendmode = inAppendmode
    End If
    '
    fnum = FreeFile()
    If inAppendmode Then
       Open infilename For Append As #fnum
    Else
       Open infilename For Output As #fnum
    End If
    isopen = True
    Print #fnum, intesto
    
Errore:
    If isopen Then Close #fnum
    If Err <> 0 Then
       If GestioneErrore(infilename) = False Then
          MsgBox "Allarm file corrupt (delete it from target folder)", vbCritical, "DP6"
       End If
    End If
End Function
