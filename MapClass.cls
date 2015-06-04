VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
' Indirizzo del blocco dati nel plc
'*******************************************************
' numero del blocco dati
Private DBNum As Long
' indirizzo prima word
Private FirstByte As Long
' numero di bytes a partire dal
Private NumBytes As Long
' dimensione max blocco dati da leggere o scrivere
Private Const MaxBytes As Integer = 200


'***********************************************************
' Variabili ausiliarie per gestione lettura e scrittura dati
'***********************************************************
' Abilitazione comunicazione con PLC
Private Enable As Boolean
' flag di simulazione abilitata
Private SimulationFlag As Boolean
'
Private ReadOnly As Boolean
Private WriteOnly As Boolean
' flag richiesta scritttura
Private WriteRequestFlag As Boolean
Private WriteInProgressFlag As Boolean

' memorizzazione indirizzo sup e inf dei dati da trasmettere
' Se max=-1 significa che non ci sono dati modificati
Private MaxModifiedData As Integer
Private MinModifiedData As Integer

' risultato di una lettura del PLC (0=ok)
Private res As Long

' tabella con i dati letti o scritti sul plc
Private PrivDB(MaxBytes / 2) As Integer
' eco dei buffer dati in lettura per rilievo dati modificati dal PLC
Private EchoPrivDB(MaxBytes / 2) As Integer
' flag segnalazione dati in lettura cambiati
Private PrivDBChangeFlag As Boolean
' forzatura dati cambiati alla prima scansione
Private FirstScanDone As Boolean

' accesso ai dati su file per simulazione PLC
Private dbs As Database
Private rstString As Recordset
Private strQuery As String
Private i As Integer

Private Sub Class_Initialize()
    Dim i As Integer
    For i = 0 To (MaxBytes / 2)
        PrivDB(i) = 0
        EchoPrivDB(i) = 0
    Next
    WriteRequestFlag = False
    WriteInProgressFlag = False
    FirstScanDone = False
    PrivDBChangeFlag = False
    MaxModifiedData = -1
    MinModifiedData = -1
    ' apre database per i/o simulato
    If Not Enable Then
        Set dbs = OpenDatabase("..\target\plc.mdb")
    End If
End Sub


' Inserimento mappa dati
Public Sub Initialize(NumeroDB As Long, ByteInizio As Long, NumeroBytes As Long, SolaLettura As Boolean, SolaScrittura As Boolean, Abilita As Boolean, Simula As Boolean)
    DBNum = NumeroDB
    FirstByte = ByteInizio
    If NumeroBytes > MaxBytes Then NumeroBytes = MaxBytes
    NumBytes = NumeroBytes
    ReadOnly = SolaLettura
    WriteOnly = SolaScrittura
    Enable = Abilita
    SimulationFlag = Simula And Not Abilita
End Sub


'************************************************************
' Funzioni di accesso ad dati di lettura-scrittura
'************************************************************
' segnalazione dati letti cambiati
Public Function DBChanged() As Boolean
    DBChanged = PrivDBChangeFlag
End Function
' reset bit dati letti cambiati
Public Sub ResetDBChanged()
    PrivDBChangeFlag = False
End Sub
' segnalazione dati letti almeno una volta
Public Function DataValid() As Boolean
    DataValid = FirstScanDone
End Function
' segnalazione simulazione abilitata
Public Function Simulation() As Boolean
    Simulation = SimulationFlag
End Function
' segnalazione scrittura in corso
Public Function Writing() As Boolean
    Writing = WriteRequestFlag Or WriteInProgressFlag
End Function


' scrittura word
Property Let Word(ByteIndex As Integer, value As Integer)
    Dim WordIndex
    WordIndex = ByteIndex / 2
    If ByteIndex >= 0 And ByteIndex < NumBytes Then
        value = kf_integer(value)
        If PrivDB(WordIndex) <> value Then
            PrivDB(WordIndex) = value
            ' accetto richiesta scrittura solo se eseguita almeno una lettura
            If WriteOnly Then
                WriteRequestFlag = True
            Else
                WriteRequestFlag = FirstScanDone
            End If
            ' memorizzazione indirizzo minimo e massimo dei dati modificati
            If MaxModifiedData < 0 Then
                MaxModifiedData = WordIndex
                MinModifiedData = WordIndex
            Else
                If MaxModifiedData < WordIndex Then MaxModifiedData = WordIndex
                If MinModifiedData > WordIndex Then MinModifiedData = WordIndex
            End If
        End If
    End If
End Property

Property Get Word(ByteIndex As Integer) As Integer
    Dim WordIndex
    WordIndex = ByteIndex / 2
    If ByteIndex >= 0 And ByteIndex < NumBytes Then
        Word = kf_integer(PrivDB(WordIndex))
    Else
        Word = 0
    End If
End Property



Property Let Bit(ByteIndex As Integer, mask As Integer, value As Boolean)
    Dim tmp As Integer
    Dim WordIndex
    WordIndex = ByteIndex / 2
    If ByteIndex >= 0 And ByteIndex < NumBytes Then
        '1) legge word
        tmp = PrivDB(WordIndex)
        'If Not Enable Then tmp = kf_integer(tmp)
        '2) modifica bit nella word
        If value Then
            tmp = tmp Or mask
        Else
            tmp = tmp And Not mask
        End If
        '3) se la word cambia di valore allora la scrive
        'If Not Enable Then tmp = kf_integer(tmp)
        If PrivDB(WordIndex) <> tmp Then
            PrivDB(WordIndex) = tmp
            ' la richiesta di scrittura viene fatta solo se almeno una lettura
            ' è stata eseguita
            If WriteOnly Then
                WriteRequestFlag = True
            Else
                WriteRequestFlag = FirstScanDone
            End If
            ' memorizzazione indirizzo minimo e massimo dei dati modificati
            If MaxModifiedData < 0 Then
                MaxModifiedData = WordIndex
                MinModifiedData = WordIndex
            Else
                If MaxModifiedData < WordIndex Then MaxModifiedData = WordIndex
                If MinModifiedData > WordIndex Then MinModifiedData = WordIndex
            End If
        End If
    End If
End Property

Property Get Bit(ByteIndex As Integer, mask As Integer) As Boolean
    Dim tmp As Integer
    Dim WordIndex
    WordIndex = ByteIndex / 2
    If ByteIndex >= 0 And ByteIndex < NumBytes Then
        tmp = PrivDB(WordIndex)
        'if not Enable Then tmp = kf_integer(tmp)
        tmp = tmp And mask
        If tmp <> 0 Then
            Bit = True
        Else
            Bit = False
        End If
    Else
        Bit = False
    End If
End Property


' legge la word senza invertire i byte per la gestione allarmi
Property Get AlarmWord(ByteIndex As Integer) As Integer
    Dim WordIndex
    WordIndex = ByteIndex / 2
    If ByteIndex >= 0 And ByteIndex < NumBytes Then
        AlarmWord = PrivDB(WordIndex)
    Else
        AlarmWord = 0
    End If
End Property


'**************************************************
'  lettura dati plc
'**************************************************
Public Function ReadAll() As Boolean
    
    If (FirstScanDone = False) Or (WriteRequestFlag = False And WriteInProgressFlag = False) Or (ReadOnly = True) Then
        ' nessuna richiesta di scrittura presente, quindi si effettua la lettura
        If Enable Then
            ' Lettura da PLC
            res = db_read(DBNum, FirstByte / 2, NumBytes / 2, EchoPrivDB(0))
        Else
'            ' Lettura da database per debug
'            Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
'            For i = FirstWord To (FirstWord + Amount - 1)
'                rstString.FindFirst ("address=" & i)
'                EchoPrivDB(i - FirstWord) = kf_integer(rstString.Fields("value"))
'            Next i
            res = 0
        End If
        If res = 0 Then
            ' segnalazione che i dati di input sono aggiornati
            For i = 0 To (NumBytes / 2) - 1
                If PrivDB(i) <> EchoPrivDB(i) Then
                    PrivDB(i) = EchoPrivDB(i)
                    PrivDBChangeFlag = True
                End If
            Next i
            ' forzatura segnalazione dati aggiornati dopo
            ' la prima scansione per assicurare che i dati
            ' vengano letti e l'immagine aggiornata anche se
            ' il DB è completamente azzerato
            If FirstScanDone = False Then
                FirstScanDone = True
                PrivDBChangeFlag = True
            End If
            ReadAll = True
       Else
            ' se c'è errore di comunicazione si riinizia il ciclo di lettura
            FirstScanDone = False
            ReadAll = False
        End If
    End If
End Function



'**************************************************
'  Scrittura dati plc
'**************************************************
Public Function WriteAll() As Boolean
    ' test se è presente una richiesta di scrittura
    If WriteRequestFlag Or WriteInProgressFlag Then
        WriteRequestFlag = False
        WriteInProgressFlag = True
        If Enable Then
            ' Scrittura su PLC
            res = db_write(DBNum, (FirstByte / 2 + MinModifiedData), (MaxModifiedData - MinModifiedData + 1), PrivDB(MinModifiedData))
        Else
            ' Scrittura su database
'            Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
'            For i = (FirstWord + MinModifiedData) To (FirstWord + MaxModifiedData)
'                rstString.FindFirst ("address=" & i)
'                rstString.Edit
'                rstString.Fields("value") = kf_integer(PrivDB(i - FirstWord))
'                rstString.Update
'            Next i
            res = 0
        End If
        'Analizza eventuali errori,
        If res = 0 Then
            PrivDBChangeFlag = True     ' forzatura aggiornamento immagine dopo scrittura
            WriteInProgressFlag = False
            MaxModifiedData = -1
            MinModifiedData = -1
            WriteAll = True
       Else
            WriteAll = False
        End If
    Else
        WriteAll = True
    End If
End Function


