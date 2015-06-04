VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecipeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Grade As String
Public Itemcode As String
Public Pieces As String
Public Weight As String
Public Destination As Integer
Public WeightPerFeet As Single

Public IDRicetta As String       ' codice della ricetta
Public TipoTubo As Integer      ' 1=rotondo 2 = sagomato
Public TuboAltezza As Double
Public TuboLarghezza As Double
Public TuboSpessore As Double
Public TuboLunghezza As Double
Public TipoPacco As Integer     ' 1=esagono 2 =quadro
Public FilaUscitaControsagoma As Integer    ' B aggiornamento automatico
Public NumeroFile As Integer                ' C aggiornamento automatico
Public NumeroTubiPacco As Integer           ' C aggiornamento automatico
Public PaccoLarghezza As Double             ' D aggiornamento automatico
Public PaccoAltezza As Double               ' E aggiornamento automatico
Public PaccoLarghezzaBaseEsagono As Double  ' F aggiornamento automatico
Public PaccoLarghezzaLatoEsagono As Double  ' G aggiornamento automatico
Public TuboPesoTeorico As Double            ' H aggiornamento automatico
Public PaccoPesoTeorico As Double           ' I aggiornamento automatico
Private PrivTubiFila(MAX_ROWS) As Integer                      ' M numero tubi per ogni fila
Public NumeroRegge As Integer
Public VelMPS As Integer                                                ' velocità MPS
Public Bypass0 As Boolean
Public Bypass1 As Boolean
Public Bypass2 As Boolean
Public Bypass3 As Boolean ' stato del bypass
Public VelVR1 As Integer
Public VelVR2 As Integer
Public VelMB1 As Integer
Public VelMB2 As Integer
Public VelMB3 As Integer
Public VelTR As Integer
Public Regg1 As Boolean
Public Regg2 As Boolean
Public TipoCalcRegge As Integer
Public Profilo As Integer
' variabili private

Private LminPaccoEsa As Double
Private LmaxPaccoEsa As Double
Private TmpFilaMax As Integer  ' per calcoli intermedi
Private HMinPacco As Double
Private HMaxPacco As Double
Private LminPacco As Double
Private LmaxPacco As Double
Private PrivQuoteRegge(MAX_STRAPS) As Double            ' N quote di reggiatura
'limiti dimensionali pacco
Property Let LarghezzaMinPaccoEsa(ByVal L_m As Double)
   LminPaccoEsa = L_m
End Property
Property Let LarghezzaMaxPaccoEsa(ByVal L_m As Double)
   LmaxPaccoEsa = L_m
End Property
Property Let LarghezzaMinPacco(ByVal L_m As Double)
   LminPacco = L_m
End Property
Property Let LarghezzaMaxPacco(ByVal L_m As Double)
   LmaxPacco = L_m
End Property
Property Let AltezzaMinPacco(ByVal L_m As Double)
   HMinPacco = L_m
End Property
Property Let AltezzaMaxPacco(ByVal L_m As Double)
   HMaxPacco = L_m
End Property
' B calcolo fila uscita controsagoma
Private Sub CalcoloFilaUscitaControsagoma()
    Dim CurRow As Integer
    Dim MaxRow As Integer
    Dim i As Integer
    For i = 0 To MAX_ROWS
        If MaxRow <= PrivTubiFila(i) Then
            MaxRow = PrivTubiFila(i)
            CurRow = i
        End If
    Next i
    FilaUscitaControsagoma = CurRow + 1
End Sub


' C calcolo numero tubi pacco (somma i tubi di ogni fila)
' poi calcola il numero di file e azzera le file successive all'ultima
Private Sub CalcoloNumeroTubiPacco()
    Dim i As Integer
    i = 1
    NumeroTubiPacco = 0
    While PrivTubiFila(i) > 0 And i < MAX_ROWS
        If PrivTubiFila(i) > 100 Then PrivTubiFila(i) = 100
        If PrivTubiFila(i) < 0 Then PrivTubiFila(i) = 0
        NumeroTubiPacco = NumeroTubiPacco + PrivTubiFila(i)
        i = i + 1
    Wend
    NumeroFile = i - 1
    For i = i To MAX_ROWS
        PrivTubiFila(i) = 0
    Next i
End Sub

' D calcolo larghezza totale pacco (m)
Public Function CalcoloPaccoLarghezza() As Double
    Dim i As Integer
    
    TmpFilaMax = 0
    For i = 1 To MAX_ROWS
        If TmpFilaMax < PrivTubiFila(i) Then TmpFilaMax = PrivTubiFila(i)
    Next i
    'calcolo larghezza pacco profili
     Select Case Profilo
    Case 0, 1
       PaccoLarghezza = TuboLarghezza * TmpFilaMax
    Case 2
       PaccoLarghezza = TuboLarghezza * (TmpFilaMax + 2)
    End Select
End Function

' E calcolo altezza totale pacco (m)
Public Function CalcoloPaccoAltezza() As Double
    Dim i As Integer
    If TipoTubo = Tubo.Tondo Then
        PaccoAltezza = TuboAltezza
        For i = 2 To MAX_ROWS
            If PrivTubiFila(i) > 0 Then
                If PrivTubiFila(i - 1) <> PrivTubiFila(i) Then
                    PaccoAltezza = PaccoAltezza + TuboAltezza * 0.866025
                Else
                    PaccoAltezza = PaccoAltezza + TuboAltezza
                End If
            End If
        Next i
    Else
        PaccoAltezza = TuboAltezza * NumeroFile
    End If
       ' calcolo altezza pacco con profili
    Select Case Profilo
    Case 1
        PaccoAltezza = (TuboAltezza + 2 * TuboSpessore) * NumeroFile
    Case 2
        PaccoAltezza = (TuboAltezza + 2 * TuboSpessore * (1 + 2 * Abs(TmpFilaMax > 2))) * NumeroFile
    End Select
End Function

' F calcolo larghezza fila base pacco esagono(m)
Public Sub CalcoloPaccoLarghezzaBaseEsagono()
    PaccoLarghezzaBaseEsagono = TuboLarghezza * (PrivTubiFila(1) - 0.42265)
End Sub

' G calcolo larghezza lato pacco esagonale(m)
Public Sub CalcoloPaccoLarghezzaLatoEsagono()
    PaccoLarghezzaLatoEsagono = ((TuboLarghezza * ((FilaUscitaControsagoma - 2#) + 1#)) + 2# * (TuboLarghezza / (2# * Sqr(3#))))
End Sub

' H calcolo peso teorico tubo (kg)
Private Sub CalcoloTuboPesoTeorico()
    On Error Resume Next
        If TipoTubo = 1 Then
            TuboPesoTeorico = (3.14159 * TuboSpessore * (TuboLarghezza - TuboSpessore)) * TuboLunghezza * Param.GetNumber("Par010_PesoSpecifico") * 1000#
        Else
            TuboPesoTeorico = (2# * TuboSpessore * (TuboLarghezza + TuboAltezza - 2 * TuboSpessore)) * TuboLunghezza * Param.GetNumber("Par010_PesoSpecifico") * 1000#
        End If
        If Err.number <> 0 Then MsgBox "Parameter ""Density"" error", vbOKOnly
    On Error GoTo 0
End Sub

' I calcolo peso teorico tubo (kg)
Private Sub CalcoloPaccoPesoTeorico()
    On Error Resume Next
        PaccoPesoTeorico = TuboPesoTeorico * NumeroTubiPacco
    On Error GoTo 0
End Sub

'M Calcolo file PACCO REGOLARE
' se pacco hex allora costruisce il pacco basandosi su fila base
' se pacco sq  allora costruisce il pacco basandosi su fila base e numerofile
Private Sub CostruisciFilePacco()
    Dim i As Integer
    If PrivTubiFila(1) < 1 Then PrivTubiFila(1) = 1         'fila base minima
    If TipoPacco = Pacco.Esagono Then
        'fila base massima con pacco hex
        If PrivTubiFila(1) > (MAX_ROWS / 2) Then PrivTubiFila(1) = (MAX_ROWS / 2)
    Else
        'fila base massima con pacco sq
        If PrivTubiFila(1) > 100 Then PrivTubiFila(1) = 100     '
    End If
    If NumeroFile < 1 Then NumeroFile = 1                   'numero file minimo
    If NumeroFile > MAX_ROWS Then NumeroFile = MAX_ROWS     'numero file massimo
    
    If TipoPacco = Pacco.Esagono Then
        If PrivTubiFila(1) > (MAX_ROWS / 2) Then PrivTubiFila(1) = (MAX_ROWS / 2)
        For i = 2 To (PrivTubiFila(1))
            PrivTubiFila(i) = PrivTubiFila(i - 1) + 1
        Next i
        For i = i To (PrivTubiFila(1) * 2 - 1)
            PrivTubiFila(i) = PrivTubiFila(i - 1) - 1
        Next i
        PrivTubiFila(i) = 0
    Else
        If TipoTubo = Tubo.Tondo Then
            ' rdsq
            For i = 2 To NumeroFile Step 2
                PrivTubiFila(i) = PrivTubiFila(1) - 1
            Next i
            For i = 3 To NumeroFile Step 2
                PrivTubiFila(i) = PrivTubiFila(1)
            Next i
        Else
            ' sqsq
            For i = 1 To NumeroFile
                PrivTubiFila(i) = PrivTubiFila(1)
            Next i
        End If
        If NumeroFile < MAX_ROWS Then PrivTubiFila(NumeroFile + 1) = 0
    End If
End Sub


Private Sub ControlloDimensioniTubo()
    If TuboLarghezza > Param.GetNumber("Par001_Tubo_LarghezzaMax") Then
        TuboLarghezza = Param.GetNumber("Par001_Tubo_LarghezzaMax")
    End If
    If TuboLarghezza < Param.GetNumber("Par002_Tubo_LarghezzaMin") Then
        TuboLarghezza = Param.GetNumber("Par002_Tubo_LarghezzaMin")
    End If
    If TuboAltezza > Param.GetNumber("Par003_Tubo_AltezzaMax") Then
        TuboAltezza = Param.GetNumber("Par003_Tubo_AltezzaMax")
    End If
    If TuboAltezza < Param.GetNumber("Par004_Tubo_AltezzaMin") Then
        TuboAltezza = Param.GetNumber("Par004_Tubo_AltezzaMin")
    End If
    If TuboSpessore > Param.GetNumber("Par005_Tubo_SpessoreMax") Then
        TuboSpessore = Param.GetNumber("Par005_Tubo_SpessoreMax")
    End If
    If TuboSpessore < Param.GetNumber("Par006_Tubo_SpessoreMin") Then
        TuboSpessore = Param.GetNumber("Par006_Tubo_SpessoreMin")
    End If
    If TuboLunghezza > Param.GetNumber("Par007_Tubo_LunghezzaMax") Then
        TuboLunghezza = Param.GetNumber("Par007_Tubo_LunghezzaMax")
    End If
    If TuboLunghezza < Param.GetNumber("Par008_Tubo_LunghezzaMin") Then
        TuboLunghezza = Param.GetNumber("Par008_Tubo_LunghezzaMin")
    End If
End Sub

'************************************************************
' Funzioni pubbliche
'************************************************************
Property Let QuotaReggia(Index As Integer, value As Double)
    If Index >= 1 And Index <= MAX_STRAPS Then
        PrivQuoteRegge(Index) = value
    End If
End Property

Property Get QuotaReggia(Index As Integer) As Double
    If Index >= 1 And Index <= MAX_STRAPS Then
        QuotaReggia = PrivQuoteRegge(Index)
    Else
        QuotaReggia = 0
    End If
End Property

Public Sub CalcoloQuoteRegge()
    Dim MaxFirstQuote, Gap, OffsetTraReggiatrici As Double
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim GapTemp As Variant
    Dim MaxReggePerLung As Integer
  
    ' limitazione NumeroReggeero regge
    If NumeroRegge > MAX_STRAPS Then NumeroRegge = MAX_STRAPS
    If NumeroRegge < 1 Then NumeroRegge = 1
    ' calcolo passo fra reggie
    
        Select Case TipoCalcRegge
        Case 0
           If PrivQuoteRegge(1) >= (TuboLunghezza / 2) Then PrivQuoteRegge(1) = Unit.m_To_Display_m(150) / 1000
           If NumeroRegge > 1 Then
                Gap = (TuboLunghezza - PrivQuoteRegge(1) * 2#) / (NumeroRegge - 1)
                ' assegnazione quote reggie
                For i = 2 To NumeroRegge
                   PrivQuoteRegge(i) = PrivQuoteRegge(i - 1) + Gap
                Next i
           End If
'        Case 1
'           =======================  calcolo Left centre =============================
'           If PrivQuoteRegge(1) >= (TuboLunghezza / 2) Then PrivQuoteRegge(1) = Unit.m_To_Display_m(150) / 1000
'           If NumeroRegge > 1 Then
'                Gap = (TuboLunghezza - PrivQuoteRegge(1) * 2#) / (NumeroRegge - 1)
'                PrivQuoteRegge(NumeroRegge) = TuboLunghezza - PrivQuoteRegge(1)
'                Gap = Param.GetNumber("Par223_OffSetRegg") / 1000
'                For i = 2 To Int(NumeroRegge / 2)
'                   PrivQuoteRegge(i) = PrivQuoteRegge(i - 1) + Gap
'                Next
'                For i = NumeroRegge - 1 To Int(NumeroRegge / 2) + 1 Step -1
'                   PrivQuoteRegge(i) = PrivQuoteRegge(i + 1) - Gap
'                Next
'
'                If (NumeroRegge And 1) > 0 Then
'                      PrivQuoteRegge(Int(NumeroRegge / 2) + 1) = (TuboLunghezza) / 2  '(TuboLunghezza - PrivQuoteRegge(1) * 2#) / 2
'                End If
'           End If
        Case 1
            Gap = Param.GetNumber("Par223_OffSetRegg") / 1000
            If NumeroRegge > 1 And Gap * (NumeroRegge - 1) <= TuboLunghezza Then
                PrivQuoteRegge(1) = 0
                For i = 2 To NumeroRegge
                   PrivQuoteRegge(i) = PrivQuoteRegge(i - 1) + Gap
                Next
                Gap = (TuboLunghezza - Gap * (NumeroRegge - 1)) / 2
                For i = 1 To NumeroRegge
                   PrivQuoteRegge(i) = PrivQuoteRegge(i) + Gap
                Next
                If (NumeroRegge And 1) > 0 And Gap * (NumeroRegge - 1) <= TuboLunghezza Then
                   PrivQuoteRegge(Int(NumeroRegge / 2) + 1) = TuboLunghezza / 2
                End If
                 
             Else
                If NumeroRegge > 1 Then
                    MaxReggePerLung = Int(TuboLunghezza / Gap)
                    PrivQuoteRegge(1) = 0
                    For i = 2 To MaxReggePerLung
                       PrivQuoteRegge(i) = PrivQuoteRegge(i - 1) + Gap
                    Next
                    Gap = (TuboLunghezza - Gap * (MaxReggePerLung - 1)) / 2
                    For i = 1 To MaxReggePerLung
                       PrivQuoteRegge(i) = PrivQuoteRegge(i) + Gap
                    Next
                    ' shift regge
                    For i = 1 To MaxReggePerLung
                       k = NumeroRegge - Int((NumeroRegge - MaxReggePerLung) / 2 + (i - 1))
                       j = MaxReggePerLung + 1 - i
                       PrivQuoteRegge(k) = _
                         PrivQuoteRegge(j)
                    Next
                    'compatta le prime regge
                    For i = Int((NumeroRegge + 1 - MaxReggePerLung) / 2) To 1 Step -1
                       k = 1 + Int((NumeroRegge + 1 - MaxReggePerLung) / 2)
                       j = i
                       GapTemp = PrivQuoteRegge(k) / (k)
                       PrivQuoteRegge(i) = PrivQuoteRegge(k) - GapTemp * j
                    Next
                    For i = Int((NumeroRegge - MaxReggePerLung) / 2) To 1 Step -1
                       k = 1 + Int((NumeroRegge + 1 - MaxReggePerLung) / 2)
                       j = NumeroRegge - i + 1
                       GapTemp = PrivQuoteRegge(k) / (Int((NumeroRegge - MaxReggePerLung) / 2) + 1)
                       PrivQuoteRegge(j) = PrivQuoteRegge(NumeroRegge - Int((NumeroRegge - MaxReggePerLung) / 2)) + GapTemp * i
                    Next
               Else
                  PrivQuoteRegge(Int(NumeroRegge / 2) + 1) = TuboLunghezza / 2
               End If
            End If
        
         Case 2
           OffsetTraReggiatrici = Param.GetNumber("Par223_OffSetRegg") / 1000 ' offset tra reggiatrici
           If PrivQuoteRegge(1) >= (TuboLunghezza / 2) Then PrivQuoteRegge(1) = 0.15
           
           If PrivQuoteRegge(1) + OffsetTraReggiatrici <= TuboLunghezza Then
                             
                      Select Case NumeroRegge
                        Case 0, 1
                        Case 2
                          PrivQuoteRegge(2) = TuboLunghezza - PrivQuoteRegge(1)
                        Case 3
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          PrivQuoteRegge(3) = TuboLunghezza - PrivQuoteRegge(1)
                          If PrivQuoteRegge(2) >= PrivQuoteRegge(3) Then PrivQuoteRegge(2) = PrivQuoteRegge(3)
                        Case 4
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(3) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = TuboLunghezza - PrivQuoteRegge(1)
                        Case 5
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(3) = (TuboLunghezza / 2)
                          PrivQuoteRegge(4) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(5) = TuboLunghezza - PrivQuoteRegge(1)
                        Case 6
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(3) = (TuboLunghezza / 2) - (OffsetTraReggiatrici / 2)
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = TuboLunghezza - PrivQuoteRegge(3)
                          PrivQuoteRegge(5) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(6) = TuboLunghezza - PrivQuoteRegge(1)
                       Case 7
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(3) = (TuboLunghezza / 2) - OffsetTraReggiatrici
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = (TuboLunghezza / 2)
                          PrivQuoteRegge(5) = TuboLunghezza - PrivQuoteRegge(3)
                          PrivQuoteRegge(6) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(7) = TuboLunghezza - PrivQuoteRegge(1)
                       Case 8
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(7) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(8) = TuboLunghezza - PrivQuoteRegge(1)
                          Gap = ((PrivQuoteRegge(7) - PrivQuoteRegge(2)) - (OffsetTraReggiatrici * 2)) / 3
                          PrivQuoteRegge(3) = Gap + PrivQuoteRegge(2)
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = PrivQuoteRegge(3) + OffsetTraReggiatrici
                          If PrivQuoteRegge(4) >= TuboLunghezza / 2 Then PrivQuoteRegge(4) = TuboLunghezza / 2
                          PrivQuoteRegge(5) = TuboLunghezza - PrivQuoteRegge(4)
                          PrivQuoteRegge(6) = TuboLunghezza - PrivQuoteRegge(3)
                       Case 9
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(8) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(9) = TuboLunghezza - PrivQuoteRegge(1)
                          PrivQuoteRegge(5) = (TuboLunghezza / 2)
                          Gap = (PrivQuoteRegge(5) - PrivQuoteRegge(2)) / 2 - (OffsetTraReggiatrici / 2) + PrivQuoteRegge(2)
                          PrivQuoteRegge(3) = Gap
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = PrivQuoteRegge(3) + OffsetTraReggiatrici
                          If PrivQuoteRegge(4) >= TuboLunghezza / 2 Then PrivQuoteRegge(4) = TuboLunghezza / 2
                          PrivQuoteRegge(6) = TuboLunghezza - PrivQuoteRegge(4)
                          PrivQuoteRegge(7) = TuboLunghezza - PrivQuoteRegge(3)
                       Case 10
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(9) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(10) = TuboLunghezza - PrivQuoteRegge(1)
                          Gap = ((PrivQuoteRegge(9) - PrivQuoteRegge(2)) - (OffsetTraReggiatrici * 3)) / 4
                          PrivQuoteRegge(3) = Gap + PrivQuoteRegge(2)
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = PrivQuoteRegge(3) + OffsetTraReggiatrici
                          If PrivQuoteRegge(4) >= TuboLunghezza / 2 Then PrivQuoteRegge(4) = TuboLunghezza / 2
                          If PrivQuoteRegge(4) < PrivQuoteRegge(3) Then PrivQuoteRegge(4) = PrivQuoteRegge(3)
                          PrivQuoteRegge(5) = Gap + PrivQuoteRegge(4)
                          If PrivQuoteRegge(5) >= TuboLunghezza / 2 Then PrivQuoteRegge(5) = TuboLunghezza / 2
                          If PrivQuoteRegge(5) < PrivQuoteRegge(4) Then PrivQuoteRegge(5) = PrivQuoteRegge(4)
                          PrivQuoteRegge(6) = TuboLunghezza - PrivQuoteRegge(5)
                          PrivQuoteRegge(7) = TuboLunghezza - PrivQuoteRegge(4)
                          PrivQuoteRegge(8) = TuboLunghezza - PrivQuoteRegge(3)
                       Case 11
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(10) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(11) = TuboLunghezza - PrivQuoteRegge(1)
                          PrivQuoteRegge(6) = (TuboLunghezza / 2)
                          Gap = (PrivQuoteRegge(6) - OffsetTraReggiatrici - PrivQuoteRegge(2)) / 2 - (OffsetTraReggiatrici / 2) + PrivQuoteRegge(2)
                          PrivQuoteRegge(3) = Gap
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = PrivQuoteRegge(3) + OffsetTraReggiatrici
                          If PrivQuoteRegge(4) >= TuboLunghezza / 2 Then PrivQuoteRegge(4) = TuboLunghezza / 2
                          If PrivQuoteRegge(4) < PrivQuoteRegge(3) Then PrivQuoteRegge(4) = PrivQuoteRegge(3)
                          PrivQuoteRegge(5) = PrivQuoteRegge(6) - OffsetTraReggiatrici
                          If PrivQuoteRegge(5) >= TuboLunghezza / 2 Then PrivQuoteRegge(5) = TuboLunghezza / 2
                          If PrivQuoteRegge(5) < PrivQuoteRegge(4) Then PrivQuoteRegge(5) = PrivQuoteRegge(4)
                          PrivQuoteRegge(7) = TuboLunghezza - PrivQuoteRegge(5)
                          PrivQuoteRegge(8) = TuboLunghezza - PrivQuoteRegge(4)
                          PrivQuoteRegge(9) = TuboLunghezza - PrivQuoteRegge(3)
                       Case 12
                          PrivQuoteRegge(2) = PrivQuoteRegge(1) + OffsetTraReggiatrici
                          If PrivQuoteRegge(2) >= TuboLunghezza / 2 Then PrivQuoteRegge(2) = TuboLunghezza / 2
                          PrivQuoteRegge(11) = TuboLunghezza - PrivQuoteRegge(2)
                          PrivQuoteRegge(12) = TuboLunghezza - PrivQuoteRegge(1)
                          Gap = ((PrivQuoteRegge(11) - PrivQuoteRegge(2)) - (OffsetTraReggiatrici * 4)) / 5
                          PrivQuoteRegge(3) = Gap + PrivQuoteRegge(2)
                          If PrivQuoteRegge(3) >= TuboLunghezza / 2 Then PrivQuoteRegge(3) = TuboLunghezza / 2
                          If PrivQuoteRegge(3) < PrivQuoteRegge(2) Then PrivQuoteRegge(3) = PrivQuoteRegge(2)
                          PrivQuoteRegge(4) = PrivQuoteRegge(3) + OffsetTraReggiatrici
                          If PrivQuoteRegge(4) >= TuboLunghezza / 2 Then PrivQuoteRegge(4) = TuboLunghezza / 2
                          If PrivQuoteRegge(4) < PrivQuoteRegge(3) Then PrivQuoteRegge(4) = PrivQuoteRegge(3)
                          PrivQuoteRegge(5) = Gap + PrivQuoteRegge(4)
                          If PrivQuoteRegge(5) >= TuboLunghezza / 2 Then PrivQuoteRegge(5) = TuboLunghezza / 2
                          If PrivQuoteRegge(5) < PrivQuoteRegge(4) Then PrivQuoteRegge(5) = PrivQuoteRegge(4)
                          PrivQuoteRegge(6) = PrivQuoteRegge(5) + OffsetTraReggiatrici
                          If PrivQuoteRegge(6) >= TuboLunghezza / 2 Then PrivQuoteRegge(6) = TuboLunghezza / 2
                          If PrivQuoteRegge(6) < PrivQuoteRegge(5) Then PrivQuoteRegge(6) = PrivQuoteRegge(5)
                          PrivQuoteRegge(7) = TuboLunghezza - PrivQuoteRegge(6)
                          PrivQuoteRegge(8) = TuboLunghezza - PrivQuoteRegge(5)
                          PrivQuoteRegge(9) = TuboLunghezza - PrivQuoteRegge(4)
                          PrivQuoteRegge(10) = TuboLunghezza - PrivQuoteRegge(3)
              
                     End Select
                   
                   Else
                   NumeroRegge = 1
                   End If
         
        End Select
    
    ' azzeramento reggie successive
    For i = NumeroRegge + 1 To MAX_STRAPS
        PrivQuoteRegge(i) = 0
    Next i
End Sub
'************************************************************
' Funzioni di accesso a dati privati
'************************************************************
Property Let TubiFila(Index As Integer, value As Integer)
    If Index >= 0 And Index < MAX_ROWS Then
        PrivTubiFila(Index) = value
    End If
End Property

Property Get TubiFila(Index As Integer) As Integer
    If Index >= 0 And Index < MAX_ROWS Then
        TubiFila = PrivTubiFila(Index)
    Else
        TubiFila = 0
    End If
End Property

Public Sub CostruisciPaccoRegolare()
    CostruisciFilePacco
    ControlloDatiPacco
End Sub

Public Sub ControlloDatiPacco()
    CalcoloNumeroTubiPacco
    ControlloDimensioniTubo
    CalcoloPaccoLarghezza
    CalcoloPaccoAltezza
    CalcoloFilaUscitaControsagoma
    CalcoloPaccoLarghezzaBaseEsagono
    CalcoloPaccoLarghezzaLatoEsagono
    CalcoloTuboPesoTeorico
    CalcoloPaccoPesoTeorico
End Sub
Public Function Cancella(ByVal IDRic As String) As Boolean
   ' Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Errore
    Cancella = False
  '  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Recipes WHERE ID='" & IDRic & "'", Connessione, , adLockOptimistic, adCmdText
     
          If .EOF = False Then
             .Delete adAffectCurrent
             Cancella = True
          End If
          .Close
          .ActiveConnection = Nothing
    End With
    Set rs = Nothing
  '  Set cn = Nothing
Exit Function

Errore:
MsgBox "Errore di connessione al database ricette", vbExclamation, "DATAPACK 6.0"
End Function
Public Function IndiceDuplicato(ByVal IDRic As String) As Boolean
  '  Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Errore
    IndiceDuplicato = False
  '  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Recipes WHERE ID='" & IDRic & "'", Connessione, , adLockReadOnly, adCmdText
     
          If .EOF = False Then IndiceDuplicato = True
          .Close
          .ActiveConnection = Nothing
    End With
    Set rs = Nothing
  '  Set cn = Nothing
Exit Function

Errore:

MsgBox "Errore di connessione al database ricette", vbExclamation, "DATAPACK 6.0"
End Function
Public Sub DownloadData(ByVal IDRic As String)
    Dim i As Integer
  '  Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
     If IDRic = "" Then Exit Sub
  '  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Recipes WHERE ID='" & IDRic & "'", Connessione, , adLockOptimistic, adCmdText
       
        ' se non trovata allora ne crea una nuova
        If .EOF = True Then
            .AddNew
            .Fields("ID") = IDRic
        End If
        
        On Error Resume Next
            ' parametri tubo
            .Fields("TipoTubo") = TipoTubo
            .Fields("Altezza") = TuboAltezza
            .Fields("Larghezza") = TuboLarghezza
            .Fields("Lunghezza") = TuboLunghezza
            .Fields("Spessore") = TuboSpessore
            .Fields("PesoTeoricoTubo") = TuboPesoTeorico
            ' parametri pacco
            .Fields("TipoPacco") = TipoPacco
            .Fields("NumeroTubi") = NumeroTubiPacco
            .Fields("NumeroFile") = NumeroFile
            .Fields("LarghezzaFilaMax") = PaccoLarghezza
            If TipoPacco = Pacco.Esagono Then
               .Fields("LarghezzaBasePacco") = PaccoLarghezzaBaseEsagono
            Else
               .Fields("LarghezzaBasePacco") = PaccoLarghezza
            End If
            .Fields("LarghezzaLatoPaccoEsagono") = PaccoLarghezzaLatoEsagono
            .Fields("AltezzaPacco") = PaccoAltezza
            .Fields("FilaUscitaControsagoma") = FilaUscitaControsagoma
            .Fields("PesoTeoricoPacco") = PaccoPesoTeorico
            For i = 1 To MAX_ROWS
                .Fields("Fila" & Format(i, "00")) = TubiFila(i)
            Next i
            ' parametri setup machina
            .Fields("R_473_64_VelMagneti") = VelMPS
            .Fields("VelVR1") = VelVR1
            .Fields("VelVR2") = VelVR2
            .Fields("VelMB1") = VelMB1
            .Fields("VelMB2") = VelMB2
            .Fields("Bypass0") = Bypass0
            .Fields("Bypass1") = Bypass1
            .Fields("Bypass2") = Bypass2
            .Fields("Bypass3") = Bypass3
            .Fields("VelTRSal") = VelTR
            .Fields("Regg1_enable") = Regg1
            .Fields("Regg2_enable") = Regg2
            .Fields("TipoCalcRegge") = TipoCalcRegge
            .Fields("Profilo") = Profilo * Abs(Param.GetBit("Par201_AbilitazioneProfili"))
            .Fields("Grade") = Grade
            .Fields("Itemcode") = Itemcode
            .Fields("Pieces") = Pieces
            .Fields("Weight") = Weight
            .Fields("WeightPerFeet") = WeightPerFeet
            ' parametri reggiatura
            .Fields("NumeroRegge") = NumeroRegge
            .Fields("Storage_destinations") = Destination
            For i = 1 To MAX_STRAPS
                .Fields("Reggia" & Format(i, "00")) = QuotaReggia(i)
            Next i
            .Update
            On Error GoTo 0
            .Close
            .ActiveConnection = Nothing
       End With
       Set rs = Nothing
     '  Set cn = Nothing
End Sub
Public Function UploadData(ByVal IDRic As String) As Boolean
    Dim i As Integer
  '  Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    
    On Error GoTo Errore
    If IDRic = "" Then IDRic = "00"
  '  cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .Open "SELECT * FROM Recipes WHERE ID='" & IDRic & "'", Connessione, , adLockReadOnly, adCmdText
      
        If .EOF = False Then
            On Error Resume Next
                ' parametri tubo
                TipoTubo = .Fields("TipoTubo")
                TuboAltezza = .Fields("Altezza")
                TuboLarghezza = .Fields("Larghezza")
                TuboLunghezza = .Fields("Lunghezza")
                TuboSpessore = .Fields("Spessore")
                TuboPesoTeorico = .Fields("PesoTeoricoTubo")
                ' parametri pacco
                TipoPacco = .Fields("TipoPacco")
                NumeroTubiPacco = .Fields("NumeroTubi")
                NumeroFile = .Fields("NumeroFile")
                PaccoLarghezza = .Fields("LarghezzaFilaMax")
                If TipoPacco = Pacco.Esagono Then
                   PaccoLarghezzaBaseEsagono = .Fields("LarghezzaBasePacco")
                Else
                   PaccoLarghezzaBaseEsagono = PaccoLarghezza
                End If
                PaccoLarghezzaLatoEsagono = .Fields("LarghezzaLatoPaccoEsagono")
                PaccoAltezza = .Fields("AltezzaPacco")
                FilaUscitaControsagoma = .Fields("FilaUscitaControsagoma")
                PaccoPesoTeorico = .Fields("PesoTeoricoPacco")
                For i = 1 To MAX_ROWS
                    TubiFila(i) = .Fields("Fila" & Format(i, "00"))
                Next i
                ' parametri setup machina
                VelMPS = .Fields("R_473_64_VelMagneti")
                VelVR1 = .Fields("VelVR1")
                VelVR2 = .Fields("VelVR2")
                VelMB1 = .Fields("VelMB1")
                VelMB2 = .Fields("VelMB2")
                VelTR = .Fields("VelTRSal")
                Bypass0 = .Fields("Bypass0")
                Bypass1 = .Fields("Bypass1")
                Bypass2 = .Fields("Bypass2")
                Bypass3 = .Fields("Bypass3")
                Regg1 = .Fields("Regg1_enable")
                Regg2 = .Fields("Regg2_enable")
                Grade = .Fields("Grade")
                WeightPerFeet = .Fields("WeightPerFeet")
                Itemcode = .Fields("Itemcode")
                Pieces = .Fields("Pieces")
                Weight = .Fields("Weight")
                TipoCalcRegge = .Fields("TipoCalcRegge")
                Destination = .Fields("Storage_destinations")
                 If IsNull(.Fields("Profilo")) = True Then
                   Profilo = 0
                Else
                   Profilo = .Fields("Profilo") * Abs(Param.GetBit("Par201_AbilitazioneProfili"))
                End If
                ' parametri reggiatura
                NumeroRegge = .Fields("NumeroRegge")
                For i = 1 To MAX_STRAPS
                    QuotaReggia(i) = .Fields("Reggia" & Format(i, "00"))
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
  '  Set cn = Nothing
Exit Function

Errore:
   MsgBox "Errore di connessione al database ricette", vbExclamation, "DATAPACK 6.0"
End Function
' controlli dimensionali pacco
Function NumMaxTubiFila() As Integer
   Dim a
   
   On Error GoTo Errore
   NumMaxTubiFila = Round(LmaxPacco / TuboLarghezza)
   a = NumMaxTubiFila * TuboLarghezza * 0.99
   If LmaxPacco < a Then NumMaxTubiFila = NumMaxTubiFila - 1
Errore:
End Function
Function NumMinTubiFila() As Integer
   Dim a
   
   On Error GoTo Errore
   NumMinTubiFila = Round(LminPacco / TuboLarghezza)
   a = NumMinTubiFila * TuboLarghezza
   If LminPacco > a Then NumMinTubiFila = NumMinTubiFila + 1
Errore:
End Function
Function NumMaxTubiFilaEsa() As Integer
      Dim a
      
       On Error GoTo Errore
       NumMaxTubiFilaEsa = Round(LmaxPaccoEsa / TuboLarghezza)
      a = NumMaxTubiFilaEsa * TuboLarghezza * 0.99
    
'      NumMaxTubiFilaEsa = Round(LmaxPacco / TuboLarghezza) / 2
'      a = NumMaxTubiFilaEsa * TuboLarghezza * 2
      If LmaxPaccoEsa < a Then NumMaxTubiFilaEsa = NumMaxTubiFilaEsa - 1
Errore:
End Function
Function NumMinTubiFilaEsa() As Integer
      Dim a
      
      On Error GoTo Errore
      NumMinTubiFilaEsa = Round(LminPaccoEsa / TuboLarghezza)
       a = NumMinTubiFilaEsa * TuboLarghezza '* 1.01
'      a = LminPacco / TuboLarghezza
'      NumMinTubiFilaEsa = Round(a) / 2
'       a = NumMinTubiFilaEsa * TuboLarghezza * 2
       If LminPaccoEsa > a Then NumMinTubiFilaEsa = NumMinTubiFilaEsa + 1
Errore:
End Function
Function NumMaxFile() As Integer
Dim a

     On Error GoTo Errore
     If TipoTubo = Tubo.Quadro Then
          NumMaxFile = Round(HMaxPacco / TuboAltezza)
          a = NumMaxFile * TuboAltezza
          If HMaxPacco < a Then NumMaxFile = NumMaxFile - 1
     Else
         NumMaxFile = HMaxPacco / (TuboAltezza * 0.866025)
         a = NumMaxFile * (TuboAltezza * 0.866025)
         If HMaxPacco < a Then NumMaxFile = NumMaxFile - 1
     End If
Errore:
End Function
Function NumMinFile() As Integer
     Dim a
     
     On Error GoTo Errore
     If TipoTubo = Tubo.Quadro Then
         NumMinFile = Round(HMinPacco / TuboAltezza)
          a = NumMinFile * TuboAltezza
          If HMinPacco > a Then NumMinFile = NumMinFile + 1
     Else
         NumMinFile = Round(HMinPacco / (TuboAltezza * 0.866025))
         a = NumMinFile * TuboAltezza * 0.866025
          If HMinPacco > a Then NumMinFile = NumMinFile + 1
     End If
Errore:
End Function



