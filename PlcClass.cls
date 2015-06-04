VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlcClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
' Indirizzo del blocco dati nel plc
'*******************************************************
' indirizzo MPI
Private Addr As Long ' Param.Number("PlcAddress")
' slot della CPU
Private Slot As Long ' Param.Number("PlcSlot")
' numero del blocco dati
Private DBNum As Long
' indirizzo prima word
Private FirstWord As Long
' numero di word a partire dalla prima
Private Amount As Long
' dimensione max blocco dati da leggere o scrivere
Private Const MaxAmount As Integer = 100


'***********************************************************
' Variabili ausiliarie per gestione lettura e scrittura dati
'***********************************************************
' flag richiesta scritttura
Private WriteRequestFlag As Boolean
Private WriteInProgressFlag As Boolean

' memorizzazione indirizzo sup e inf dei dati da trasmettere
' Se max=-1 significa che non ci sono dati modificati
Private MaxModifiedData As Integer
Private MinModifiedData As Integer

' risultato di una lettura del PLC (0=ok)
Private res As Long
Private adr_table As String

' Abilitazione comunicazione con PLC
Private Enable As Boolean
' flag di simulazione abilitata
Private SimulationFlag As Boolean
' flag che contraddistingue la mappa dati principale (per gestione errori di comunicazione)
Private Master As Boolean
' tabella con i dati letti o scritti sul plc
Private PrivDB(MaxAmount) As Integer
' eco dei buffer dati in lettura per rilievo dati modificati dal PLC
Private EchoPrivDB(MaxAmount) As Integer
' flag segnalazione dati in lettura cambiati
Private PrivDBChangeFlag As Boolean
' forzatura dati cambiati alla prima scansione
Private FirstScanDone As Boolean

' accesso ai dati su file per simulazione PLC
Private dbs As Database
Private rstString As Recordset
Private strQuery As String
Private i As Integer


#If S7Plc Then
    '*******************************************************
    ' funzione importate del driver PRODAVE
    '*******************************************************
    ' conversione dati
    Private Declare Function kf_integer Lib "komfort.dll" (ByVal nr As Integer) As Integer
'    Private Declare Function float_to_gp Lib "komfort.dll" (yyy%, xxx%)
'    Private Declare Function gp_to_float Lib "komfort.dll" (xxx%, yyy%)
    ' inizializzazione
    Private Declare Function load_tool Lib "w95_s7m.dll" (ByVal nr As Long, ByVal dev$, ByVal adr$) As Long
    ' fine della comunicazione
    ' NB: è pericoloso terminare il programma senza aver chiamato questa funzione
    Private Declare Function unload_tool Lib "w95_s7m.dll" () As Long
    ' lettura
    Private Declare Function db_read Lib "w95_s7m.dll" (ByVal db As Long, ByVal dw As Long, anz As Long, value%) As Long
    ' scrittura
    Private Declare Function db_write Lib "w95_s7m.dll" (ByVal db As Long, ByVal dw As Long, anz As Long, value%) As Long
    ' ripristino errore di comunicazione
    Private Declare Function new_ss Lib "w95_s7m.dll" (ByVal nr As Long) As Long
    '*******************************************************
    ' fine funzione importate del driver PRODAVE
    '*******************************************************
    
#Else   ' non S7
    '*******************************************************
    ' funzioni fittizie non utilizzare le librerie siemens
    ' quando si utilizzano altri Plc
    '*******************************************************
    Private Function kf_integer(ByVal nr As Integer) As Integer
        kf_integer = nr
    End Function
    Private Function load_tool(ByVal nr As Long, ByVal dev$, ByVal adr$) As Long
        load_tool = 0
    End Function
    Private Function unload_tool() As Long
        unload_tool = 0
    End Function
    Private Function db_read(ByVal db As Long, ByVal dw As Long, anz As Long, value%) As Long
        db_read = 0
    End Function
    Private Function db_write(ByVal db As Long, ByVal dw As Long, anz As Long, value%) As Long
        db_write = 0
    End Function
    Private Function new_ss(ByVal nr As Long) As Long
        new_ss = 0
    End Function
    '*******************************************************
    ' fine funzioni fittizie
    '*******************************************************
#End If

Private Sub Class_Initialize()
    Dim i As Integer
    For i = 0 To MaxAmount
        PrivDB(i) = 0
        EchoPrivDB(i) = 0
    Next
    Enable = Param.Bit("PlcEnable")
    SimulationFlag = Param.Bit("PlcSimulation") And (Not Enable)
    Addr = Param.Number("PlcAddress")
    Slot = Param.Number("PlcSlot")
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
Public Sub Initialize(m As Boolean, db As Long, fw As Long, am As Long)
    ' si chiama la funzione "Hide" per forzare la chiamata
    ' della load_form prima di inizializzare i dati
    Master = m
    DBNum = db
    FirstWord = fw
    If am > MaxAmount Then am = MaxAmount
    Amount = am
End Sub

Private Sub Class_Terminate()
    If Master Then res = unload_tool()
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
' set DB in sola scrittura
Public Sub WriteOnly()
    ' evita di dover leggere il DB per poterlo scrivere
    FirstScanDone = True
End Sub
' segnalazione simulazione abilitata
Public Function Simulation() As Boolean
    Simulation = SimulationFlag
End Function
' segnalazione scrittura in corso
Public Function Writing() As Boolean
    Writing = WriteRequestFlag Or WriteInProgressFlag
End Function


' scrittura word
Property Let Word(Index As Integer, value As Integer)
    If Index >= 0 And Index < Amount Then
        value = kf_integer(value)
        If PrivDB(Index) <> value Then
            PrivDB(Index) = value
            ' accetto richiesta scrittura solo se eseguita almeno una lettura
            WriteRequestFlag = FirstScanDone
            ' memorizzazione indirizzo minimo e massimo dei dati modificati
            If MaxModifiedData < 0 Then
                MaxModifiedData = Index
                MinModifiedData = Index
            Else
                If MaxModifiedData < Index Then MaxModifiedData = Index
                If MinModifiedData > Index Then MinModifiedData = Index
            End If
        End If
    End If
End Property

Property Get Word(Index As Integer) As Integer
    If Index >= 0 And Index < Amount Then
        Word = kf_integer(PrivDB(Index))
    Else
        Word = 0
    End If
End Property


'' scrittura real
'Property Let Real(Index As Integer, value As Double)
'    Dim Matrice(2) As Integer
'    Dim prec1, prec2 As Integer
'    Matrice(0) = value
'    If Index >= 0 And Index < Amount Then
'        ' memorizza valori precedenti
'        prec1 = PrivDB(Index)
'        prec2 = PrivDB(Index + 1)
'        ' assegna nuovo valore
'        float_to_gp value, PrivDB(Index)
'        If (prec1 <> PrivDB(Index)) Or (prec2 <> PrivDB(Index + 1)) Then
'            ' accetto richiesta scrittura solo se eseguita almeno una lettura
'            WriteRequestFlag = FirstScanDone
'            ' memorizzazione indirizzo minimo e massimo dei dati modificati
'            If MaxModifiedData < 0 Then
'                MaxModifiedData = Index + 1
'                MinModifiedData = Index
'            Else
'                If MaxModifiedData <= Index Then MaxModifiedData = Index + 1
'                If MinModifiedData > Index Then MinModifiedData = Index
'            End If
'        End If
'    End If
'End Property



Property Let Bit(Index As Integer, mask As Integer, value As Boolean)
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        '1) legge word
        tmp = PrivDB(Index)
        'If Not Enable Then tmp = kf_integer(tmp)
        '2) modifica bit nella word
        If value Then
            tmp = tmp Or mask
        Else
            tmp = tmp And Not mask
        End If
        '3) se la word cambia di valore allora la scrive
        'If Not Enable Then tmp = kf_integer(tmp)
        If PrivDB(Index) <> tmp Then
            PrivDB(Index) = tmp
            WriteRequestFlag = FirstScanDone
            ' memorizzazione indirizzo minimo e massimo dei dati modificati
            If MaxModifiedData < 0 Then
                MaxModifiedData = Index
                MinModifiedData = Index
            Else
                If MaxModifiedData < Index Then MaxModifiedData = Index
                If MinModifiedData > Index Then MinModifiedData = Index
            End If
        End If
    End If
End Property

Property Get Bit(Index As Integer, mask As Integer) As Boolean
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        tmp = PrivDB(Index)
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
Property Get AlarmWord(Index As Integer) As Integer
    If Index >= 0 And Index < Amount Then
        AlarmWord = PrivDB(Index)
    Else
        AlarmWord = 0
    End If
End Property


'**************************************************
'  lettura dati plc
'**************************************************
Public Sub ReadAll()
    ' apertura della connessione
    ' (FirstScanDone=False dopo ogni errore di comunicazione)
    ' NB: il DB master deve essere ReadWrite, perchè il ripristino della
    ' comunicazione è effettuato solo da questa funzione di lettura e non
    ' da quella di scrittura
    If Master And Enable And FirstScanDone = False Then
        res = unload_tool()
        adr_table = Chr(Addr) + Chr(0) + Chr(Slot) + Chr(0) + Chr(0)
        res = load_tool(1, "S7ONLINE", adr_table)
        If res = 0 Then
            AlarmForm.PcAlarm(PlcDDEFault) = False
        Else
            AlarmForm.PcAlarm(PlcDDEFault) = True
        End If
    End If

    ' segnalazione allarme comunicazione non abilitata
    If Master Then
        AlarmForm.PcAlarm(PlcNotEnable) = Not Enable
    End If
    
    If (FirstScanDone = False) Or (WriteRequestFlag = False And WriteInProgressFlag = False) Then
        ' nessuna richiesta di scrittura presente, quindi si effettua la lettura
        If Enable Then
            ' Lettura da PLC
            res = db_read(DBNum, FirstWord, Amount, EchoPrivDB(0))
        Else
            ' Lettura da database per debug
            Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
            For i = FirstWord To (FirstWord + Amount - 1)
                rstString.FindFirst ("address=" & i)
                EchoPrivDB(i - FirstWord) = kf_integer(rstString.Fields("value"))
            Next i
            res = 0
        End If
        If res = 0 Then
            ' segnalazione che i dati di input sono aggiornati
            For i = 0 To Amount - 1
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
            If Master Then AlarmForm.PcAlarm(PlcCommErr) = False
        Else
            ' se c'è errore di comunicazione si riinizia il ciclo di lettura
            If Master Then AlarmForm.PcAlarm(PlcCommErr) = True
            FirstScanDone = False
        End If
    End If
End Sub

' arresto ciclo lettura per risparmiare tempo
Public Sub ReadStop()
    'Questa funzione fa qualcosa solo nel protocollo AB
    ' (arresto ciclo lettura automatica via DDE dopo la lettura iniziale)
End Sub


'**************************************************
'  Scrittura dati plc
'**************************************************
Public Sub WriteAll()
    ' test se è presente una richiesta di scrittura
    If WriteRequestFlag Or WriteInProgressFlag Then
        
        WriteRequestFlag = False
        WriteInProgressFlag = True
        If Enable Then
            ' Scrittura su PLC
            res = db_write(DBNum, (FirstWord + MinModifiedData), (MaxModifiedData - MinModifiedData + 1), PrivDB(MinModifiedData))
        Else
            ' Scrittura su database
            Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
            For i = (FirstWord + MinModifiedData) To (FirstWord + MaxModifiedData)
                rstString.FindFirst ("address=" & i)
                rstString.Edit
                rstString.Fields("value") = kf_integer(PrivDB(i - FirstWord))
                rstString.Update
            Next i
        End If
        'Analizza eventuali errori,
        If res = 0 Then
            PrivDBChangeFlag = True     ' forzatura aggiornamento immagine dopo scrittura
            WriteInProgressFlag = False
            MaxModifiedData = -1
            MinModifiedData = -1
        End If
    End If
End Sub


