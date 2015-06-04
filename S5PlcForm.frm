VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form S5PlcForm 
   Caption         =   "S5Plc"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSCommLib.MSComm PlcMSComm 
      Left            =   600
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   1024
      ParitySetting   =   2
      StopBits        =   2
      InputMode       =   1
   End
End
Attribute VB_Name = "S5PlcForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
' Indirizzo del PLC
'*******************************************************
' Il numero della porta seriale connessa al PLC
' è il parametro "PlcSlot"

' Si consiglia di mettere un tempo di update di 250 ms

' dimensione max blocco dati da leggere o scrivere
Private Const Amount As Integer = 100


'*******************************************************
' Variabili da inizializzare nel costruttore
'*******************************************************
Private DBNum As Long

Private ReadAmount As Long
Private ReadFirstWord As Long

Private WriteAmount As Long
Private WriteFirstWord As Long

Private ReadWriteAmount As Long
Private ReadWriteFirstWord As Long

Private WriteRequestFlag As Boolean
Private WriteInProgressFlag As Boolean
Private InOutWriteRequestFlag As Boolean
Private InOutWriteInProgressFlag As Boolean

Private res As Integer ' risultato di una lettura del PLC (0=ok)

' Abilitazione comunicazione con PLC
Private Enable As Boolean
'Private EnableEcho As Boolean ' per rilevare fronti di abilitazione
Public Simulation As Boolean ' flag di simulazione abilitata
Private Master As Boolean
Private DelayCounter As Integer
Private ValidAddress As Boolean
' flag=false durante il ciclo di lettura zona Read-only
' flag=true durante il ciclo di lettura zona Read-Write
Private ReadOnlyDone As Boolean


'***************************************************************
' Variabile di stato
' 0....99       Lettura indirizzo DB
' 100...199     Lettura DB
' 200...299     Scrittura DB
'***************************************************************
Private State As Integer
' buffer TX RX
Private InBuffer() As Byte
Private TmpInBuffer() As Byte       ' per ricezione DB
Private ReceivedBytes As Integer    ' per ricezione DB
Private CommandMode As Boolean      ' per ricezione DB
Private Offset As Integer           ' per ricezione DB
Private OutBuffer() As Byte
' bytes predefiniti
Private Const NUL As Byte = 0
Private Const SOH As Byte = 1
Private Const STX As Byte = 2
Private Const ETX As Byte = 3
Private Const EOT As Byte = 4
Private Const ACK As Byte = 6
Private Const DLE As Byte = &H10
Private Const NAK As Byte = &H15
'indirizzi della dw0
Private MSBAddress As Byte
Private LSBAddress As Byte
'indirizzi della zona read
Private MSBFirstReadAddress As Byte
Private LSBFirstReadAddress As Byte
Private MSBLastReadAddress As Byte
Private LSBLastReadAddress As Byte
'indirizzi della zona read-write
Private MSBFirstReadWriteAddress As Byte
Private LSBFirstReadWriteAddress As Byte
Private MSBLastReadWriteAddress As Byte
Private LSBLastReadWriteAddress As Byte
'indirizzi della zona write
Private MSBFirstWriteAddress As Byte
Private LSBFirstWriteAddress As Byte


Private Const Timeout As Integer = 10

Private PrivReadDB(Amount) As Integer
Private PrivWriteDB(Amount) As Integer
Private PrivReadWriteDB(Amount) As Integer

' eco dei buffer dati in lettura per rilievo dati modificati
' dal PLC al termine del ciclo di lettura
Private EchoPrivReadDB(Amount) As Integer
Private EchoPrivReadWriteDB(Amount) As Integer
' flag segnalazione dati cambiati
Private PrivReadDBChangeFlag As Boolean
Private PrivReadWriteDBChangeFlag As Boolean
Private FirstScanDone As Boolean ' forzatura dati cambiati alla prima scansione

' accesso ai dati su file per simulazione PLC
Private dbs As Database
Private rstString As Recordset
Private strQuery As String
Private i As Integer

' conversione dati
'Private Declare Function kf_integer Lib "komfort.dll" (ByVal nr As Integer) As Integer



' Inserimento mappa dati
Public Sub Initialize(m As Boolean, db As Long, rfw As Long, ra As Long, _
    wfw As Long, wa As Long, rwfw As Long, rwa As Long)
    ' si chiama la funzione show per forzare la chiamata
    ' della load_form prima di inizializzare i dati
    Me.Show
    Me.Caption = "S5 DB" & db
    Master = m
    DBNum = db
    ReadFirstWord = rfw
    ReadAmount = ra
    WriteFirstWord = wfw
    WriteAmount = wa
    ReadWriteFirstWord = rwfw
    ReadWriteAmount = rwa
    
    PrivReadDBChangeFlag = False
    PrivReadWriteDBChangeFlag = False
    FirstScanDone = False
    
    ' apre database per i/o simulato
    Set dbs = OpenDatabase("..\target\plc.mdb")
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To Amount
        PrivReadDB(i) = 0
        PrivWriteDB(i) = 0
        PrivReadWriteDB(i) = 0
    Next
    Master = False
    Enable = False
    DelayCounter = 0
    WriteRequestFlag = False
    WriteInProgressFlag = False
    InOutWriteRequestFlag = False
    InOutWriteInProgressFlag = False
    DBNum = 1
    If Param.Bit("PlcEnable") Then
        Enable = True
        State = 10
    Else
        Enable = False
        State = 0
    End If
    Simulation = Param.Bit("PlcSimulation") And (Not Enable)
    On Error Resume Next
        PlcMSComm.CommPort = Param.Number("PlcSlot")
        If Err.Number <> 0 Then MsgBox "Parameter ""PlcSlot"" error", vbOKOnly
    On Error GoTo 0
    ValidAddress = False
    

End Sub


'**************************************************
' Funzione per impedire la chiusura del form da parte
' dell'utente
'**************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormControlMenu Then
        Cancel = False
    Else
        Cancel = True
    End If
End Sub

'************************************************************
' Funzioni di accesso ad dati di lettura-scrittura
'************************************************************
Property Get ReadWriteDBChanged() As Boolean
    ReadWriteDBChanged = PrivReadWriteDBChangeFlag
End Property

Public Sub ResetReadWriteDBChanged()
    PrivReadWriteDBChangeFlag = False
End Sub

Property Get ReadWriteDataValid() As Boolean
    ReadWriteDataValid = FirstScanDone
End Property

Property Let ReadWriteWord(Index As Integer, value As Integer)
    If Index >= 0 And Index < Amount Then
        PrivReadWriteDB(Index) = value
        InOutWriteRequestFlag = FirstScanDone
    End If
End Property

Property Get ReadWriteWord(Index As Integer) As Integer
    If Index >= 0 And Index < Amount Then
        ReadWriteWord = PrivReadWriteDB(Index)
    Else
        ReadWriteWord = 0
    End If
End Property

Property Let ReadWriteBit(Index As Integer, mask As Integer, value As Boolean)
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        tmp = PrivReadWriteDB(Index)
        If value Then
            tmp = tmp Or mask
        Else
            tmp = tmp And Not mask
        End If
        PrivReadWriteDB(Index) = tmp
        InOutWriteRequestFlag = FirstScanDone
    End If
End Property

Property Get ReadWriteBit(Index As Integer, mask As Integer) As Boolean
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        tmp = PrivReadWriteDB(Index)
        tmp = tmp And mask
        If tmp <> 0 Then
            ReadWriteBit = True
        Else
            ReadWriteBit = False
        End If
    Else
        ReadWriteBit = False
    End If
End Property


'************************************************************
' Funzioni di accesso ad dati di scrittura
'************************************************************

Property Let WriteWord(Index As Integer, value As Integer)
    If Index >= 0 And Index < Amount Then
        PrivWriteDB(Index) = value
        WriteRequestFlag = True
    End If
End Property

Property Get WriteWord(Index As Integer) As Integer
    If Index >= 0 And Index < Amount Then
        WriteWord = PrivWriteDB(Index)
    Else
        WriteWord = 0
    End If
End Property

Property Let WriteBit(Index As Integer, mask As Integer, value As Boolean)
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        tmp = PrivWriteDB(Index)
        If value Then
            tmp = tmp Or mask
        Else
            tmp = tmp And Not mask
        End If
        PrivWriteDB(Index) = tmp
        WriteRequestFlag = True
    End If
End Property


Property Get WriteBit(Index As Integer, mask As Integer) As Boolean
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        tmp = PrivWriteDB(Index)
        tmp = tmp And mask
        If tmp <> 0 Then
            WriteBit = True
        Else
            WriteBit = False
        End If
    Else
        WriteBit = False
    End If
End Property

' per visualizzare finestra scrittura tabelle in corso
Public Function OutWriteInProgress() As Boolean
    OutWriteInProgress = WriteRequestFlag Or WriteInProgressFlag
End Function


'************************************************************
' Funzioni di accesso ad dati di lettura
'************************************************************

Property Get ReadDBChanged() As Boolean
    ReadDBChanged = PrivReadDBChangeFlag
End Property

Public Sub ResetReadDBChanged()
    PrivReadDBChangeFlag = False
End Sub

Property Get ReadWord(Index As Integer) As Integer
    If Index >= 0 And Index < Amount Then
        ReadWord = PrivReadDB(Index)
    Else
        ReadWord = 0
    End If
End Property
' funzione a disposizione del simulatore di plc
Property Let ReadWord(Index As Integer, value As Integer)
    If Simulation Then
        If Index >= 0 And Index < Amount And Not Enable Then
            EchoPrivReadDB(Index) = value
        End If
    End If
End Property
' legge la word senza invertire i byte per la gestione allarmi
Property Get ReadAlarmWord(Index As Integer) As Integer
    If Index >= 0 And Index < Amount Then
        ReadAlarmWord = PrivReadDB(Index)
    Else
        ReadAlarmWord = 0
    End If
End Property

Property Get ReadBit(Index As Integer, mask As Integer) As Boolean
    Dim tmp As Integer
    If Index >= 0 And Index < Amount Then
        tmp = PrivReadDB(Index)
        tmp = tmp And mask
        If tmp <> 0 Then
            ReadBit = True
        Else
            ReadBit = False
        End If
    Else
        ReadBit = False
    End If
End Property


' Aggiorna la comunicazione con il plc
'***************************************************************
' Variabile di stato
' 0....99       Lettura indirizzo DB
' 100...199     Lettura DB
' 200...299     Scrittura DB
'***************************************************************
Public Sub Dfa()
    ' controllo timeout tra richiami dfa
    'CurrentTime = Time()
    ' if(((CurrentTime-ElapsedTime)>400) && (stato>=110)  && (Param.PlcEnable==TRUE)){
    '    // Questo cambio di stato non funziona sul PC epson
    '    // probabilmente perchè il CurrentTime non viene aggiornato in
    '    // modo regolare.
    '
    '        // stato=105;
    '  }//if
    '  ElapsedTime=CurrentTime;

    Select Case State
            
        Case 0   ' simulazione lettura PLC
            If DelayCounter > 0 Then
                DelayCounter = DelayCounter - 1
            Else
                ' test se è presente una richiesta di scrittura per la zona read-write
                If InOutWriteRequestFlag = False And InOutWriteInProgressFlag = False Then
                    ' nessuna richiesta presente, quindi si effettua la lettura
                    ReadInputOutputArea
                Else
                    InOutWriteRequestFlag = False
                    InOutWriteInProgressFlag = True
                    If WriteInputOutputArea Then InOutWriteInProgressFlag = False
                End If
                ' test se c'è una richiesta di scrittura per la zona write only
                If WriteRequestFlag = True Or WriteInProgressFlag = True Then
                    WriteRequestFlag = False
                    WriteInProgressFlag = True
                    If WriteOutputArea Then WriteInProgressFlag = False
                End If
                ' lettura area read only
                ReadInputArea
                ' Si abilita la pausa
                DelayCounter = 5
                MainMDIForm.PcAlarm(PlcNotEnable) = True
            End If


        Case 10  ' init seriale
            On Error Resume Next
                If Not PlcMSComm.PortOpen Then PlcMSComm.PortOpen = True
                If Err.Number = 0 Then
                    If ValidAddress Then
                        State = 100 ' indirizzi già validi, passa alla lettura del DB
                    Else
                        State = 12  ' prosegue con lettura indirizzi
                    End If
                End If
            On Error GoTo 0

' ACQUISIZIONE INDIRIZZI

        Case 12   ' inizio acquisizione indirizzi
            ' pulizia buffer di ricezione
            InBuffer = PlcMSComm.Input
            ' trasmissione carattere STX.
            ReDim OutBuffer(0)
            OutBuffer(0) = STX
            PlcMSComm.Output = OutBuffer
            DelayCounter = 0
            State = 14

        Case 14  ' // attesa fine ricezione
            If PlcMSComm.InBufferCount >= 2 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(1) = ACK Then
                    ' dati ok trasmette  0x1A
                    ReDim OutBuffer(0)
                    OutBuffer(0) = &H1A
                    PlcMSComm.Output = OutBuffer
                    State = 16
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300


        Case 16
            If PlcMSComm.InBufferCount >= 1 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(0) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 18
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 18
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = ETX Then
                    ReDim OutBuffer(20)
                    i = 0
                    OutBuffer(i) = DLE: i = i + 1
                    OutBuffer(i) = ACK: i = i + 1
                    OutBuffer(i) = &H1: i = i + 1
                    OutBuffer(i) = DBNum: i = i + 1
                    'raddoppio del DLE se il dato e' &H10
                    If DBNum = &H10 Then OutBuffer(i) = DLE: i = i + 1
                    OutBuffer(i) = DLE: i = i + 1
                    OutBuffer(i) = EOT: i = i + 1
                    ReDim Preserve OutBuffer(i - 1)
                    PlcMSComm.Output = OutBuffer
                    State = 20
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 20
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 22
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 22
            If PlcMSComm.InBufferCount >= 4 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(0) = NUL Then
                    MSBAddress = InBuffer(1)
                    If MSBAddress = DLE Then
                        LSBAddress = InBuffer(3)
                    Else
                        LSBAddress = InBuffer(2)
                    End If
                    'calcolo degli altri indirizzi
                    AddressUpdate

                    ' trasmissione ACK
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    ValidAddress = True
                    State = 24
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 24     ' pausa necessaria dopo lettura indirizzi
            DelayCounter = DelayCounter + 1
            If DelayCounter >= 10 Then State = 100

' FINE ACQUISIZIONE INDIRIZZI

' LETTURA DATI
        Case 100    ' Inizio lettura DB
            ' pulizia buffer di ricezione
            InBuffer = PlcMSComm.Input
            ' trasmissione carattere STX.
            ReDim OutBuffer(0)
            OutBuffer(0) = STX
            PlcMSComm.Output = OutBuffer
            DelayCounter = 0
            State = 102

        Case 102
            If PlcMSComm.InBufferCount >= 2 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(1) = ACK Then
                    ' dati ok trasmette  0x04
                    ReDim OutBuffer(0)
                    OutBuffer(0) = &H4
                    PlcMSComm.Output = OutBuffer
                    State = 104
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 104
            If PlcMSComm.InBufferCount >= 1 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(0) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 106
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 106
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = ETX Then
                    ReDim OutBuffer(20)
                    i = 0
                    OutBuffer(i) = DLE: i = i + 1
                    OutBuffer(i) = ACK: i = i + 1
                    If Not ReadOnlyDone Then
                        OutBuffer(i) = MSBFirstReadAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                        OutBuffer(i) = LSBFirstReadAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                        OutBuffer(i) = MSBLastReadAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                        OutBuffer(i) = LSBLastReadAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                    Else
                        OutBuffer(i) = MSBFirstReadWriteAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                        OutBuffer(i) = LSBFirstReadWriteAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                        OutBuffer(i) = MSBLastReadWriteAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                        OutBuffer(i) = LSBLastReadWriteAddress: i = i + 1
                        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
                    End If
                    OutBuffer(i) = DLE: i = i + 1
                    OutBuffer(i) = EOT: i = i + 1
                    ReDim Preserve OutBuffer(i - 1)
                    PlcMSComm.Output = OutBuffer
                    State = 108
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300


        Case 108
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 110
                    i = 0
                    DelayCounter = 0
                    ReDim TmpInBuffer(0)
                    CommandMode = False
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300



        Case 110 ' ricezione DB
            ' memorizzazione dei caratteri ricevuti nel ReadBuffer
            If PlcMSComm.InBufferCount > 0 Then
                ' se ci sono caratteri nuovi ricevuti allora si leggono
                InBuffer = PlcMSComm.Input
                ReceivedBytes = UBound(InBuffer) + 1
                ' si crea lo spazio per il trasferimento
                ReDim Preserve TmpInBuffer(ReceivedBytes + UBound(TmpInBuffer))
                ' si trasferiscono i dati analizzandoli per capire se
                ' il DB è stato ricevuto completamente
                Offset = i
                While (i - Offset) < ReceivedBytes
                    TmpInBuffer(i) = InBuffer(i - Offset)
                    If CommandMode Then
                        If TmpInBuffer(i) = ETX Then
                            ' ricezione terminata
                            ReceivedBytes = 0   ' per forzare uscita da ciclo while
                            ReDim OutBuffer(1)
                            OutBuffer(0) = DLE
                            OutBuffer(1) = ACK
                            PlcMSComm.Output = OutBuffer
                            DelayCounter = 0
                            State = 112
                        Else
                            CommandMode = False
                        End If
                    Else
                        If TmpInBuffer(i) = DLE Then CommandMode = True
                    End If
                    i = i + 1
                Wend
            End If
            DelayCounter = DelayCounter + 1
            If (DelayCounter > Timeout * 4) Then State = 300

        Case 112
            If PlcMSComm.InBufferCount >= 1 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(0) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 114
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 114
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = ETX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 116
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300



        Case 116: 'trasferimento dati  TmpInBuffer -> PrivReadDB
            If Not ReadOnlyDone Then
                ' terminato il primo ciclo (Lettura zona read-only)
                ' traferisce i dati da eco a privDB read-write controllando se
                ' ci sono state variazioni nei dati
                DBTransfer
                ReadOnlyDone = True
                ' ritorna all'inizio del ciclo di lettura dati per
                ' leggere la zona read-write
                State = 100
            Else
                ' terminato il secondo ciclo (lettura zona read-write)
                If InOutWriteRequestFlag Or InOutWriteInProgressFlag Then
                    ' richiesta scrittura in corso, passa alla scrittura
                    ' senza trasferire l'eco nel DB read-write
                    InOutWriteInProgressFlag = True
                    InOutWriteRequestFlag = False
                    State = 200
                Else
                    ' traferisce i dati da eco a privDB read-write controllando se
                    ' ci sono state variazioni nei dati
                    DBTransfer
                    ' forzatura segnalazione dati aggiornati dopo
                    ' la prima scansione per assicurare che i dati
                    ' vengano letti e l'immagine aggiornata anche se
                    ' il DB è completamente azzerato
                    If FirstScanDone = False Then
                        FirstScanDone = True
                        PrivReadWriteDBChangeFlag = True
                    End If
                    ' infine si testa se ci sono richieste di scrittura
                    If WriteRequestFlag Or WriteInProgressFlag Then
                        WriteInProgressFlag = True
                        WriteRequestFlag = False
                        State = 200
                    Else
                        'lettura terminata e non ci sono  richieste di scrittura
                        State = 299
                    End If
                End If
                ReadOnlyDone = False
            End If
        
' FINE LETTURA DATI


' SCRITTURA DATI
        Case 200
            ' pulizia buffer di ricezione
            InBuffer = PlcMSComm.Input
            ' trasmissione carattere STX.
            ReDim OutBuffer(0)
            OutBuffer(0) = STX
            PlcMSComm.Output = OutBuffer
            DelayCounter = 0
            State = 202
            
        Case 202
            If PlcMSComm.InBufferCount >= 2 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(1) = ACK Then
                    ' dati ok trasmette  0x03
                    ReDim OutBuffer(0)
                    OutBuffer(0) = &H3
                    PlcMSComm.Output = OutBuffer
                    State = 204
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300
            

        Case 204
            If PlcMSComm.InBufferCount >= 1 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(0) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 206
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300


        Case 206    ' trasmette il DB
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = ETX Then
                    OutBufferLoad
                    PlcMSComm.Output = OutBuffer
                    State = 208
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

        Case 208
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = STX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    State = 210
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > (Timeout * 4) Then State = 300

        Case 210
            If PlcMSComm.InBufferCount >= 3 Then
                InBuffer = PlcMSComm.Input
                If InBuffer(2) = ETX Then
                    ReDim OutBuffer(1)
                    OutBuffer(0) = DLE
                    OutBuffer(1) = ACK
                    PlcMSComm.Output = OutBuffer
                    If WriteInProgressFlag Then
                        ' terminata la scrittura zona Write-Only
                        WriteInProgressFlag = False
                    Else
                        ' terminata la scrittura zona Read-Write
                        InOutWriteInProgressFlag = False
                    End If
                    State = 299
                    DelayCounter = 0
                End If
            End If
            DelayCounter = DelayCounter + 1
            If DelayCounter > Timeout Then State = 300

' FINE SCRITTURA DATI

        Case 299  ' chiusura ciclo senza errori
            PlcMSComm.PortOpen = False
            DelayCounter = 0
            MainMDIForm.PcAlarm(PlcCommErr) = False
            State = 10  ' ritorna all'inizio
        
        Case 300    ' timeout
            PlcMSComm.PortOpen = False
            MainMDIForm.PcAlarm(PlcCommErr) = True
            ValidAddress = False
            DelayCounter = 0
            State = 302

        Case 302    ' pausa dopo timeout
            DelayCounter = DelayCounter + 1
            If DelayCounter > 10 Then State = 10


        Case Else
            State = 0
    End Select

End Sub



'**************************************************
'  lettura zona read only (simulata e solo per il primo ciclo)
'**************************************************
Private Function ReadInputArea() As Boolean
    Static FirstReadDone As Boolean
    If (Not Simulation) Or (Not FirstReadDone) Then
        ' Lettura da database per debug
        Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
        For i = ReadFirstWord To (ReadFirstWord + ReadAmount - 1)
            rstString.FindFirst ("address=" & i)
            EchoPrivReadDB(i - ReadFirstWord) = rstString.Fields("value")
        Next i
        FirstReadDone = True
    End If
    ' trasferimento dati da buffer provvisorio a finale
    ' con controllo cambio dati
    ReadInputArea = True
    For i = 0 To ReadAmount - 1
        If PrivReadDB(i) <> EchoPrivReadDB(i) Then
            PrivReadDB(i) = EchoPrivReadDB(i)
            PrivReadDBChangeFlag = True
        End If
    Next i
End Function

'**************************************************
'  lettura zona read write (simulata)
'**************************************************
Private Function ReadInputOutputArea() As Boolean
    Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
    For i = ReadWriteFirstWord To (ReadWriteFirstWord + ReadWriteAmount - 1)
        rstString.FindFirst ("address=" & i)
        EchoPrivReadWriteDB(i - ReadWriteFirstWord) = rstString.Fields("value")
    Next i
    ' segnalazione che i dati di input sono aggiornati
    ReadInputOutputArea = True
    For i = 0 To ReadWriteAmount - 1
        If PrivReadWriteDB(i) <> EchoPrivReadWriteDB(i) Then
            PrivReadWriteDB(i) = EchoPrivReadWriteDB(i)
            PrivReadWriteDBChangeFlag = True
        End If
    Next i
    ' forzatura segnalazione dati aggiornati dopo
    ' la prima scansione per assicurare che i dati
    ' vengano letti e l'immagine aggiornata anche se
    ' il DB è completamente azzerato
    If FirstScanDone = False Then
        FirstScanDone = True
        PrivReadWriteDBChangeFlag = True
    End If
End Function


'**************************************************
'  Scrittura zona write only (simulata)
'**************************************************
Private Function WriteOutputArea() As Boolean
    Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
    For i = WriteFirstWord To (WriteFirstWord + WriteAmount - 1)
        rstString.FindFirst ("address=" & i)
        rstString.Edit
        rstString.Fields("value") = PrivWriteDB(i - WriteFirstWord)
        rstString.Update
    Next i
    WriteOutputArea = True
End Function


'**************************************************
'  Scrittura zona read write (simulata)
'**************************************************
Private Function WriteInputOutputArea() As Boolean
    Set rstString = dbs.OpenRecordset("DB" & DBNum, dbOpenDynaset)
    For i = ReadWriteFirstWord To (ReadWriteFirstWord + ReadWriteAmount - 1)
        rstString.FindFirst ("address=" & i)
        rstString.Edit
        rstString.Fields("value") = PrivReadWriteDB(i - ReadWriteFirstWord)
        rstString.Update
    Next i
    WriteInputOutputArea = True
End Function



' calcolo indirizzi di tutte le zone del DB a partire da
' MSBAddress e LSBAddress della DW0
Private Sub AddressUpdate()
    Dim DW0Address As Long
    Dim TmpAddress As Long
    
    ' calcolo dell'indirizzo della dw0 del DB
    DW0Address = MSBAddress
    DW0Address = DW0Address * 256
    DW0Address = DW0Address Or LSBAddress
    
    ' calcolo dell'indirizzo iniziale della zona read only
    TmpAddress = DW0Address + (ReadFirstWord * 2)
    LSBFirstReadAddress = TmpAddress And &HFF
    MSBFirstReadAddress = (TmpAddress \ 256) And &HFF
    ' calcolo dell'indirizzo finale della zona read only
    TmpAddress = DW0Address + (ReadFirstWord + ReadAmount) * 2 - 1
    LSBLastReadAddress = TmpAddress And &HFF
    MSBLastReadAddress = (TmpAddress \ 256) And &HFF
    
    ' calcolo dell'indirizzo iniziale della zona read-write
    TmpAddress = DW0Address + (ReadWriteFirstWord * 2)
    LSBFirstReadWriteAddress = TmpAddress And &HFF
    MSBFirstReadWriteAddress = (TmpAddress \ 256) And &HFF
    ' calcolo dell'indirizzo finale della zona read-write
    TmpAddress = DW0Address + (ReadWriteFirstWord + ReadWriteAmount) * 2 - 1
    LSBLastReadWriteAddress = TmpAddress And &HFF
    MSBLastReadWriteAddress = (TmpAddress \ 256) And &HFF
    
    ' calcolo dell'indirizzo iniziale della zona write-only
    TmpAddress = DW0Address + (WriteFirstWord * 2)
    LSBFirstWriteAddress = TmpAddress And &HFF
    MSBFirstWriteAddress = (TmpAddress \ 256) And &HFF

End Sub
            
' legge dati da TmpInBuffer e li trasferisce in EchoPriv togliendo i DLE
' duplicati e poi trasferisce in Priv rilevando le variazioni dei dati letti
'NB: si basa sul flag ReadOnlyDone per determinare dove vanno trasferiti i dati
Private Sub DBTransfer()
    Dim LongData As Long
    Dim j As Integer
    Dim MaxI As Integer
    'i=conteggio word    j=conteggio bytes
    If Not ReadOnlyDone Then
        MaxI = ReadAmount
    Else
        MaxI = ReadWriteAmount
    End If
    j = 5: i = 0
    On Error Resume Next ' maschera errori di dimensione TmpInBuffer
        While i < MaxI
            ' load MSB
            If TmpInBuffer(j) <> DLE Then
                LongData = TmpInBuffer(j): j = j + 1
            Else
                LongData = TmpInBuffer(j + 1): j = j + 2
            End If
            LongData = LongData * 256
            ' load LSB
            If TmpInBuffer(j) <> DLE Then
                LongData = LongData Or TmpInBuffer(j): j = j + 1
            Else
                LongData = LongData Or TmpInBuffer(j + 1): j = j + 2
            End If
            ' se il bit(16)=1 allora mette a 1 anche i 16 bit più
            ' significativi per evitare errori nel trasferimento del dato
            ' dalla variabile a 32 bit alla variabile a 16 bit
            If LongData > &H7FFF Then LongData = LongData Or &HFFFF0000
            ' trasferisce la word
            If Not ReadOnlyDone Then
                EchoPrivReadDB(i) = LongData
                If EchoPrivReadDB(i) <> PrivReadDB(i) Then
                    PrivReadDB(i) = EchoPrivReadDB(i)
                    PrivReadDBChangeFlag = True
                End If
            Else
                EchoPrivReadWriteDB(i) = LongData
                If EchoPrivReadWriteDB(i) <> PrivReadWriteDB(i) Then
                    PrivReadWriteDB(i) = EchoPrivReadWriteDB(i)
                    PrivReadWriteDBChangeFlag = True
                End If
            End If
            i = i + 1
        Wend
    On Error GoTo 0
End Sub

' prepara buffer di trasmissione per scrittura DB
Private Sub OutBufferLoad()
    Dim LongData As Long
    Dim j As Integer
    'i=conteggio word    j=conteggio bytes
    ReDim OutBuffer(512)
    i = 0: j = 0
    OutBuffer(i) = DLE: i = i + 1
    OutBuffer(i) = ACK: i = i + 1
    If WriteInProgressFlag Then
        ' scrittura zona Write-Only
        OutBuffer(i) = MSBFirstWriteAddress: i = i + 1
        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
        OutBuffer(i) = LSBFirstWriteAddress: i = i + 1
        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
        While j < WriteAmount
            ' MSB data
            LongData = PrivWriteDB(j)
            LongData = LongData And &HFFFF&
            LongData = LongData \ 256
            OutBuffer(i) = LongData: i = i + 1
            If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
            ' LSB data
            LongData = PrivWriteDB(j)
            LongData = LongData And &HFF&
            OutBuffer(i) = LongData: i = i + 1
            If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
            j = j + 1
        Wend
    Else
        ' scrittura zona Read-Write
        OutBuffer(i) = MSBFirstReadWriteAddress: i = i + 1
        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
        OutBuffer(i) = LSBFirstReadWriteAddress: i = i + 1
        If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
        While j < ReadWriteAmount
            ' MSB data
            LongData = PrivReadWriteDB(j)
            LongData = LongData And &HFFFF&
            LongData = LongData \ 256
            OutBuffer(i) = LongData: i = i + 1
            If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
            ' LSB data
            LongData = PrivReadWriteDB(j)
            LongData = LongData And &HFF&
            OutBuffer(i) = LongData: i = i + 1
            If OutBuffer(i - 1) = DLE Then OutBuffer(i) = DLE: i = i + 1
            j = j + 1
        Wend
    End If
    OutBuffer(i) = DLE: i = i + 1
    OutBuffer(i) = EOT: i = i + 1
    ReDim Preserve OutBuffer(i - 1)
End Sub

