VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BundleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BundleId As Long  ' codice del pacco nell'archivio
Public Hex As Boolean       ' pacco esagono
Public RoundTube As Boolean ' tubo rotondo
Public Regular As Boolean   ' pacco regolare
Private PrivTubesRow(MAX_ROWS) As Integer   ' numero tubi per ogni fila
Private PrivTubes As Integer

''************************************************************
'' Funzioni di accesso ad dati privati
''************************************************************
Property Let TubesRow(Index As Integer, value As Integer)
    If Index >= 0 And Index < MAX_ROWS Then
        PrivTubesRow(Index) = value
    End If
    ' ricalcola numero totale tubi ad ogni modifica della composizione del pacco
    CalcTubes
End Property

Property Get TubesRow(Index As Integer) As Integer
    If Index >= 0 And Index < MAX_ROWS Then
        TubesRow = PrivTubesRow(Index)
    Else
        TubesRow = 0
    End If
End Property

Property Get Tubes() As Integer
    Tubes = PrivTubes
End Property

' calcolo numero file pacco
Public Function Rows() As Integer
    Dim i As Integer
    Rows = MAX_ROWS
    For i = MAX_ROWS To 0 Step -1
        If PrivTubesRow(i) = 0 Then Rows = i
    Next i
End Function

' calcolo numero tubi pacco
Private Sub CalcTubes()
    Dim i As Integer
    PrivTubes = 0
    For i = 0 To MAX_ROWS
        PrivTubes = PrivTubes + PrivTubesRow(i)
    Next i
End Sub

' calcolo larghezza totale pacco (m)
Public Function Width(TubeWidth As Double) As Double
    Dim i, MaxRow As Integer ' per calcoli intermedi
    MaxRow = 0
    For i = 0 To MAX_ROWS
        If MaxRow < PrivTubesRow(i) Then MaxRow = PrivTubesRow(i)
    Next i
    Width = TubeWidth * MaxRow
    If Hex Then
        Width = Width + TubeWidth * 0.155
    End If
End Function
' calcolo altezza totale pacco (m)
Public Function Height(TubeHeight As Double) As Double
    Dim i As Integer
    If RoundTube Then
        'Height = TubeHeight * (Rows * 0.866025 + 0.133975)
        Height = TubeHeight
        For i = 1 To MAX_ROWS
            If PrivTubesRow(i) > 0 Then
                If PrivTubesRow(i - 1) <> PrivTubesRow(i) Then
                    Height = Height + TubeHeight * 0.866025
                Else
                    Height = Height + TubeHeight
                End If
            End If
        Next i

    Else
        Height = (TubeHeight) * Rows
    End If
End Function

' calcolo larghezza fila base pacco (m)
Public Function Base(TubeWidth As Double) As Double
    If RoundTube Then
        If Hex Then
            Base = TubeWidth * (PrivTubesRow(0) - 0.42265)
        Else
            Base = TubeWidth * PrivTubesRow(0)
        End If
    Else
        Base = TubeWidth * PrivTubesRow(0)
    End If
End Function

' calcolo larghezza lato pacco (m)
Public Function Side(TubeHeight As Double) As Double
    If RoundTube Then
        If Hex Then
            Side = ((RowTemplate() - 1) - 0.42265) * TubeHeight
        Else
            Side = TubeHeight * (Rows * 0.866025 + 0.133975)
        End If
    Else
        Side = (TubeHeight) * Rows
    End If
End Function

' calcolo fila uscita controsagoma
Public Function RowTemplate() As Integer
    Dim CurRow, MaxRow, i As Integer ' per calcoli intermedi
    MaxRow = 0
    CurRow = 0
    For i = 0 To MAX_ROWS
        If MaxRow <= PrivTubesRow(i) Then
            MaxRow = PrivTubesRow(i)
            CurRow = i
        End If
    Next i
    RowTemplate = CurRow + 2
End Function
'
'' costruisce un pacco regolare basandosi sulla fila base
'' e sul numero file, da richiamare quando cambia la fila base
'Private Sub BaseChange()
'    Dim i As Integer
'    If PrivTubesRow(0) < 2 Then PrivTubesRow(0) = 2
'    If PrivTubesRow(0) > MAX_ROWS Then PrivTubesRow(0) = MAX_ROWS
'    If Hex Then
'        If PrivTubesRow(0) > (MAX_ROWS / 2) Then PrivTubesRow(0) = (MAX_ROWS / 2)
'        For i = 1 To (PrivTubesRow(0) - 1)
'            PrivTubesRow(i) = PrivTubesRow(i - 1) + 1
'        Next i
'        For i = PrivTubesRow(0) To (PrivTubesRow(0) * 2 - 2)
'            PrivTubesRow(i) = PrivTubesRow(i - 1) - 1
'        Next i
'        For i = (PrivTubesRow(0) * 2 - 1) To MAX_ROWS
'            PrivTubesRow(i) = 0
'        Next i
'    Else
'        If RoundTube Then
'            ' rdsq
'            If PrivTubesRow(0) = 1 Then
'                PrivTubesRow(1) = 0
'                PrivTubesRow(2) = 0
'            End If
'            For i = 1 To MAX_ROWS Step 2
'                If PrivTubesRow(i) > 0 Then PrivTubesRow(i) = PrivTubesRow(0) - 1
'            Next i
'            For i = 2 To MAX_ROWS Step 2
'                If PrivTubesRow(i) > 0 Then PrivTubesRow(i) = PrivTubesRow(0)
'            Next i
'        Else
'            ' sqsq
'            For i = 1 To MAX_ROWS
'                If PrivTubesRow(i) > 0 Then PrivTubesRow(i) = PrivTubesRow(0)
'            Next i
'        End If
'    End If
'End Sub
'
'' costruisce un pacco quadro regolare, da richiamare quando
'' cambia il numero di file nel pacco quadro
'Public Sub RowsChange(RowNumber As Integer)
'    Dim i As Integer
'    If Not Hex Then
'        If RoundTube Then
'            ' rdsq
'            For i = 1 To MAX_ROWS Step 2
'                If i < RowNumber Then
'                    PrivTubesRow(i) = PrivTubesRow(0) - 1
'                Else
'                    PrivTubesRow(i) = 0
'                End If
'            Next i
'            For i = 2 To MAX_ROWS Step 2
'                If i < RowNumber Then
'                    PrivTubesRow(i) = PrivTubesRow(0)
'                Else
'                    PrivTubesRow(i) = 0
'                End If
'            Next i
'        Else
'            ' sqsq
'            For i = 1 To MAX_ROWS
'                If i < RowNumber Then
'                    PrivTubesRow(i) = PrivTubesRow(0)
'                Else
'                    PrivTubesRow(i) = 0
'                End If
'            Next i
'        End If
'    End If
'End Sub
'
'' costruisce un pacco regolare, da richiamare quando
'' cambia il numero il tipo di pacco o tubo o fila base
'Public Sub AutomaticRows()
'    Dim i As Integer
'    ' controlla flag tipo pacco e tubo
'    If Not RoundTube Then Hex = False
'    ' controlla dimensioni tubo
'    If RoundTube Then TubeWidth = TubeHeight
'    ' ricostruisce le file del pacco
'    BaseChange
'End Sub
'


' controllo file pacco speciale
Public Sub DataCheck()
    Dim i As Integer
    ' controlla flag tipo pacco e tubo
    If Not RoundTube Then Hex = False
    ' controlla file
    For i = 0 To MAX_ROWS
        If PrivTubesRow(i) > MAX_ROWS Then PrivTubesRow(i) = MAX_ROWS
        If PrivTubesRow(i) < 0 Then PrivTubesRow(i) = 0
    Next i
    For i = Rows To MAX_ROWS
        PrivTubesRow(i) = 0
    Next i
End Sub






