VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Da riportare su 1975.pc :
' coefficente libbre/kg
' ..
' Costanti per rappresentazione in unità decimali
' dati per gestire i kg
Private Const MetricKgString As String = " (Kg)"
Private Const MetricKgCoeff As Double = 1#
' dati per gestire i mm
Private Const MetricMMString As String = " (mm)"
Private Const MetricMMCoeff As Double = 1000#       ' mm/m
Private Const MetricMMFormatString As String = "###0"
Private Const MetricMMFormatString0 As String = "##0.0"
Private Const MetricMMFormatString00 As String = "#0.00"

Private Const MetricMMSliderCoeff As Double = 1#
Private Const MetricMMSliderCoeff0 As Double = 10#
Private Const MetricMMSliderCoeff00 As Double = 100#
' dati per gestire i m
Private Const MetricMSpeedString As String = " (m/min)"
Private Const MetricMString As String = " (m)"

Private Const MetricMCoeff As Double = 1#          '  m/m
Private Const MetricMFormatString As String = "###0"
Private Const MetricMFormatString0 As String = "##0.0"
Private Const MetricMSliderCoeff As Double = 1#
Private Const MetricMSliderCoeff0 As Double = 10#


' Costanti per rappresentazione in unità anglosassoni
' dati per sostituire i kg con i pound
Private Const AngloKgString As String = " (lb)"
Private Const AngloKgCoeff As Double = 2.205 '3.7854  ' lb/kg

' dati per sostituire i mm con inch ( un decimale in piu')
Private Const AngloMMString As String = " (in)"
Private Const AngloMMCoeff As Double = 39.37007874   ' in/m
Private Const AngloMMFormatString As String = "##0.0"
Private Const AngloMMFormatString0 As String = "###0.00"
Private Const AngloMMFormatString00 As String = "#0.000"
Private Const AngloMMSliderCoeff As Double = 10#
Private Const AngloMMSliderCoeff0 As Double = 100#
Private Const AngloMMSliderCoeff00 As Double = 1000#

' dati per sostituire i m con i ft ( stessi decimali)
Private Const AngloMSpeedString As String = " (ft/min)"
Private Const AngloMString As String = " (ft)"
Private Const AngloMCoeff As Double = 3.28083989    ' ft/m
Private Const AngloMFormatString As String = "###0"
Private Const AngloMFormatString0 As String = "##0.0"
Private Const AngloMSliderCoeff As Double = 1#
Private Const AngloMSliderCoeff0 As Double = 10#


' Variabili interne

Private privMMString As String
Private MMCoeff As Double
Private MMFormatString As String
Private MMFormatString0 As String
Private MMFormatString00 As String
Private KgCoeff As Double
Private privKgString As String
Private MMSliderCoeff As Double
Private MMSliderCoeff0 As Double
Private MMSliderCoeff00 As Double
Private PrivMSpeedString As String
Private PrivMString As String
Private MCoeff As Double
Private MFormatString As String
Private MFormatString0 As String
Private MSliderCoeff As Double
Private MSliderCoeff0 As Double

Private Sub Class_Initialize()
    If Param.GetBit("Par101_MisureMetriche") Then
        KgCoeff = AngloKgCoeff
        privKgString = AngloKgString
        privMMString = AngloMMString
        MMCoeff = AngloMMCoeff
        MMFormatString = AngloMMFormatString
        MMFormatString0 = AngloMMFormatString0
        MMFormatString00 = AngloMMFormatString00
        MMSliderCoeff = AngloMMSliderCoeff
        MMSliderCoeff0 = AngloMMSliderCoeff0
        MMSliderCoeff00 = AngloMMSliderCoeff00
        PrivMSpeedString = AngloMSpeedString
        PrivMString = AngloMString
        MCoeff = AngloMCoeff
        MFormatString = AngloMFormatString
        MFormatString0 = AngloMFormatString0
        MSliderCoeff = AngloMSliderCoeff
        MSliderCoeff0 = AngloMSliderCoeff0
    Else
        KgCoeff = MetricKgCoeff
        privKgString = MetricKgString
        privMMString = MetricMMString
        MMCoeff = MetricMMCoeff
        MMFormatString = MetricMMFormatString
        MMFormatString0 = MetricMMFormatString0
        MMFormatString00 = MetricMMFormatString00
        MMSliderCoeff = MetricMMSliderCoeff
        MMSliderCoeff0 = MetricMMSliderCoeff0
        MMSliderCoeff00 = MetricMMSliderCoeff00
        PrivMSpeedString = MetricMSpeedString
        PrivMString = MetricMString
        MCoeff = MetricMCoeff
        MFormatString = MetricMFormatString
        MFormatString0 = MetricMFormatString0
        MSliderCoeff = MetricMSliderCoeff
        MSliderCoeff0 = MetricMSliderCoeff0
    End If
End Sub

' funzioni di accesso in sola lettura alle stringhe
' con le unità di misura
Property Get mmString() As String
    mmString = privMMString
End Property
Property Get MtMinString() As String
    MtMinString = PrivMSpeedString
End Property
Property Get mString() As String
    mString = PrivMString
End Property
Property Get KgString() As String
    KgString = privKgString
End Property

'conversione per visualizzare

' funzioni di conversione da Double(m) a stringa per
' display in mm con 0, 1 o 2 decimali
Public Function m_To_Display_mm(value As Double) As String
    m_To_Display_mm = Format(value * MMCoeff, MMFormatString)
End Function
Public Function m_To_Display_mm0(value As Double) As String
    m_To_Display_mm0 = Format(value * MMCoeff, MMFormatString0)
End Function
Public Function m_To_Display_mm00(value As Double) As String
    m_To_Display_mm00 = Format(value * MMCoeff, MMFormatString00)
End Function

' funzioni di conversione da Double(m) a intero
' per slider in mm con 0, 1 o 2 decimali
Public Function m_To_Slider_mm(value As Double) As Long
    m_To_Slider_mm = value * MMCoeff * MMSliderCoeff
End Function
Public Function m_To_Slider_mm0(value As Double) As Long
    m_To_Slider_mm0 = value * MMCoeff * MMSliderCoeff0
End Function
Public Function m_To_Slider_mm00(value As Double) As Long
    m_To_Slider_mm00 = value * MMCoeff * MMSliderCoeff00
End Function

' funzioni di conversione da intero
' (slider con 0, 1 o 2 decimali) a Double(m)
Public Function Slider_mm_To_m(value As Long) As Double
    Slider_mm_To_m = value / (MMCoeff * MMSliderCoeff)
End Function
Public Function Slider_mm0_To_m(value As Long) As Double
    Slider_mm0_To_m = value / (MMCoeff * MMSliderCoeff0)
End Function
Public Function Slider_mm00_To_m(value As Long) As Double
    Slider_mm00_To_m = value / (MMCoeff * MMSliderCoeff00)
End Function

' conversione da introduzione

' funzioni di conversione da stringa(mm) a Double(m)
Public Function Display_mm_To_m(value As String, Prev As Double) As Double
    Display_mm_To_m = Prev
    On Error Resume Next
    Display_mm_To_m = CDbl(value) / MMCoeff
End Function

'***********************************************************
'        m <---> m
'***********************************************************
' funzioni di conversione da Double(m) a stringa per
' display in m con 0 decimali
Public Function m_To_Display_m(value As Double) As String
    m_To_Display_m = Format(value * MCoeff, MFormatString)
End Function
' funzioni di conversione da Double(m) a intero
' per slider in m con 0 decimali
Public Function m_To_Slider_m(value As Double) As Long
    m_To_Slider_m = value * MCoeff * MSliderCoeff
End Function
' funzioni di conversione da intero
' (slider con 0 decimali) a Double(m)
Public Function Slider_m_To_m(value As Long) As Double
    Slider_m_To_m = value / (MCoeff * MSliderCoeff)
End Function
' funzioni di conversione da stringa(m) a Double(m)
Public Function Display_m_To_m(value As String, Prev As Double) As Double
    Display_m_To_m = Prev
    On Error Resume Next
    Display_m_To_m = CDbl(value) / MCoeff
End Function


'***********************************************************
'        Kg <---> Kg
'***********************************************************
' funzioni di conversione da Double(Kg) a stringa per
' display in Kg o pound con 0 decimali
Public Function kg_To_Display_kg(value As Double) As String
'    If value > 0 Then
        kg_To_Display_kg = Format(value * KgCoeff, MFormatString)
'    Else
'        kg_To_Display_kg = "---"
'    End If
End Function

'***********************************************************
'        Int <---> Int
'***********************************************************
' funzioni di conversione da stringa(int) a int
Public Function Display_int_To_int(value As String, Prev As Integer) As Integer
    Display_int_To_int = Prev
    On Error Resume Next
    Display_int_To_int = CInt(value)
End Function

'====================== NUOVE FUNZIONI ==========================

Public Function OUT_Kg(value As Double) As String
     OUT_Kg = Format(value * KgCoeff, MFormatString) & privKgString
End Function
Public Function IN_m_OUT_Dec_mm(value As Double) As String
    IN_m_OUT_Dec_mm = Format(value * MMCoeff, MMFormatString0)
End Function

