VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmKernel 
   BorderStyle     =   0  'None
   Caption         =   "frmKERNEL"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   Icon            =   "Kernel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSAdodcLib.Adodc AdoGrade 
      Height          =   525
      Left            =   120
      Top             =   4260
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   926
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoGrade"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer TSimula 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   3120
   End
   Begin VB.Timer TimerHook 
      Enabled         =   0   'False
      Interval        =   19
      Left            =   120
      Top             =   2670
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   150
      Top             =   2220
   End
   Begin MSAdodcLib.Adodc AdoLogAllarmi 
      Height          =   495
      Left            =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Kernel.frx":08CA
      OLEDBString     =   $"Kernel.frx":0959
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AllarmiMultiLingua"
      Caption         =   "AdoLogAllarmi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoLogPacco 
      Height          =   495
      Left            =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdoLogPacco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdoAllarmi 
      Height          =   495
      Left            =   120
      Top             =   3720
      Visible         =   0   'False
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Kernel.frx":09E8
      OLEDBString     =   $"Kernel.frx":0A77
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "AllarmiMultiLingua"
      Caption         =   "AdoAllarmi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmKernel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const EWX_SHUTDOWN As Long = 1

'======================= indirizzi PLC siemens S7 ======================
' nome del server (visibile con opc scout)
Const S7_Server = "OPC.SimaticNET"
' collegamento assegnato dal server opc PB 5.3
'Const S7_nome_collegamento = "S7:[S7_connection_name1|VFD1|CP_L2_1:]DB"
'Const S7_nome_collegamentoI_O = "S7:[S7_connection_name1|VFD1|CP_L2_1:]"
' collegamento con PB6.0-6.1
'Const S7_nome_collegamento = "S7:[S7 connection_1]DB"
'Const S7_nome_collegamentoI_O = "S7:[S7 connection_1]"

'================================================================
Const PB_MAX = 800

'Definizione dell'elenco di pagine disponibili

Public Enum Pagina
    PagOrdini = 1
    PagStorico = 2
    PagLayout = 3
    PagPacco = 4
    PagRegge = 5
    PagPesa = 6
    Pagallarmi = 7
    PagSmusso = 8
    PagFiletto = 9
    PagWb = 10
    PagTaglio = 11
    PagEntrata = 12
    PagService = 13
    PagLavaggio = 14
End Enum

'pagina con focus
Public ContaChiusurasys As Integer
Public AllResettati As Boolean
Public NumAllAttivi As Long
Public NumMsgAttivi As Long
Public ServerOpcOn As Boolean
Public SimulaON As Boolean
Public ServerDisattivo As Boolean
Private BitPesa As Boolean
Private m_Allarme As Boolean
Private m_PaginaCorrente As Pagina
Private stato As Integer
Private SBstep As String

Public PaccoArchiviato As Boolean
Public CodOrdineCorrente As COrdineAttuale
Public AllarmeON As Boolean
Public Errore As Boolean
Public PulAllarmiPremuto As Boolean
Public IDOrdineCorrente As String
Public DescrOrdineCorrente As String

'*************************************************
' Mappa degli allarmi interni
'*************************************************
Private PcAlarmWord As Integer
Private Boot As Boolean

Public Enum TPcAlarm
    PlcCommErr = &H1        ' allarme interno 1
    PlcNotEnable = &H2      ' allarme interno 2
    PlcDDEFault = &H4       ' allarme interno 3
    Alarm4 = &H8            ' allarme interno 4
    PrinterCommErr = &H10   ' allarme interno 5
    PrinterFault = &H20     ' allarme interno 6
    PrinterCommPort = &H40  ' allarme interno 7
    WeightCommErr = &H80    ' allarme interno 8
    WeightFault = &H100     ' allarme interno 9
    WeightCommPort = &H200  ' allarme interno 10
    WeightReqFault = &H400  ' allarme interno 11
    ErrCompattDataBase = &H800 ' allarme interno 12
    ErrRipristDataBase = &H1000 ' allarme interno 13
    Alarm14 = &H2000        ' allarme interno 14
    Alarm15 = &H4000        ' allarme interno 15
    TrakingError = &H8000   ' allarme interno 16
End Enum

' ordine e ricetta per download dati nelle varie zone del traking

Private KernelRecipe As RecipeClass
' pubblici

Public KernelOrder As OrderClass
Public StrapEnableFlag As Boolean
Public LabelEnableFlag As Boolean

'assegnazione della pagina corrente e cambio pagina (visualizzazione form)

Property Let PaginaCorrente(ByVal nuovaPagina As Pagina)
    m_PaginaCorrente = nuovaPagina
    
    Select Case m_PaginaCorrente
     Case Pagina.PagEntrata
                DB450.Refresh
                frmEntrata.ZOrder (0)
                frmEntrata.WindowState = vbMaximized
                frmEntrata.Show
        Case Pagina.Pagallarmi
             If PulAllarmiPremuto = False Then
               AlarmForm.CheckDB400.value = 1
               AlarmForm.CheckDB410.value = 1
               AlarmForm.CheckDB411.value = 1
               AlarmForm.CheckDB412.value = 1
               AlarmForm.CheckDB413.value = 1
               AlarmForm.CheckDB414.value = 1
               AlarmForm.CheckDB415.value = 1
               AlarmForm.CheckDB416.value = 1
               AlarmForm.CheckDB417.value = 1
               AlarmForm.CheckDB418.value = 1
               AlarmForm.CheckDB419.value = 1
               AlarmForm.CheckDB420.value = 1
               AlarmForm.CheckDB422.value = 1
               AlarmForm.CheckDB424.value = 1
               AlarmForm.CheckDB425.value = 1
               AlarmForm.CheckDB426.value = 1
             End If
                AlarmForm.ZOrder (0)
                AlarmForm.WindowState = vbMaximized
                AlarmForm.Show
        Case Pagina.PagFiletto
                DB460.Refresh
                FilettoForm.SSTab1.Tab = 0
                FilettoForm.ZOrder (0)
                FilettoForm.WindowState = vbMaximized
                FilettoForm.Show
        Case Pagina.PagLayout
                CommandForm.ZOrder (0)
                CommandForm.WindowState = vbMaximized
                CommandForm.Show
        Case Pagina.PagOrdini
                OrdersForm.ZOrder (0)
                OrdersForm.WindowState = vbMaximized
                OrdersForm.Show
        Case Pagina.PagPacco
                BundleForm.ZOrder (0)
                BundleForm.WindowState = vbMaximized
                OrdersForm.SSTab1.Tab = 1
                BundleForm.Show
        Case Pagina.PagPesa
                frmStoccaggio.ZOrder (0)
                frmStoccaggio.WindowState = vbMaximized
                frmStoccaggio.Show
        Case Pagina.PagRegge
                DB480.Refresh
                FormRegge.ZOrder (0)
                FormRegge.WindowState = vbMaximized
                FormRegge.Show
        Case Pagina.PagService
                Param.ZOrder (0)
                Param.WindowState = vbMaximized
                Param.Frame4.ZOrder
                Param.Barra21.ZOrder
                Param.Show
        Case Pagina.PagSmusso
                DB460.Refresh
                frmSmussatrice.SSTab1.Tab = 0
                frmSmussatrice.ZOrder (0)
                frmSmussatrice.WindowState = vbMaximized
                frmSmussatrice.Show
        Case Pagina.PagStorico
                BundlesLogForm.ZOrder (0)
                BundlesLogForm.WindowState = vbMaximized
                BundlesLogForm.Show
        Case Pagina.PagWb
                DB460.Refresh
                FormWB.ZOrder (0)
                FormWB.WindowState = vbMaximized
                FormWB.Show
        Case Pagina.PagLavaggio
                DB465.Refresh
                FormLavaggio.ZOrder (0)
                FormLavaggio.WindowState = vbMaximized
                FormLavaggio.Show
    End Select
End Property
Property Get PaginaCorrente() As Pagina
   PaginaCorrente = m_PaginaCorrente
End Property

Property Let PcAlarm(mask As TPcAlarm, value As Boolean)
    If value Then
        PcAlarmWord = PcAlarmWord Or mask
    Else
        PcAlarmWord = PcAlarmWord And Not mask
    End If
End Property

Property Get PcAlarm(mask As TPcAlarm) As Boolean
    If (PcAlarmWord And mask) <> 0 Then
        PcAlarm = True
    Else
        PcAlarm = False
    End If
End Property

Private Sub Form_Load()
    Dim a() As String
    Dim i As Integer
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim TmpValore As Variant
    
    Timer1.Enabled = False
    Timer1.Interval = 400
    TimerHook.Enabled = False
    TimerHook.Interval = 100
    TSimula.Enabled = False
    TSimula.Interval = 250
    
    '========= controlla il database ===================
    RipristinaDatabase
    '============================================
    '============================================
    ' Collegamento con MSRM
    '============================================
    On Error Resume Next
'    Text1.LinkTopic = "LabelProc|Form1"
'    Text1.LinkItem = "Text2"
'    Text1.LinkMode = vbLinkNone
'    Text1.LinkMode = vbLinkManual
    '============================================
    Boot = False
    ErrDBPiccolo = False
    Show
    WindowState = vbMinimized
    
    'inizializza lo stato degli allarmi
    
    AllarmeON = False
 
    '============================================================================
    'COSTRUZIONE OGGETTI DICHIARATI IN GlobalDataKernel
    '============================================================================
    
    ' Stato delle zone e allarmi
    
    Set DB400 = New DBClass    ' stato linea (RO)
    DB400.DB_ID = "400"
    Set DB410 = New DBClass    ' stato zona entrata (RO)
    DB410.DB_ID = "410"
    Set DB411 = New DBClass    ' stato zona accumulo (RO)
    DB411.DB_ID = "411"
    'Set DB412 = New DBClass    ' stato zona entrata taglio (RO)
    'DB412.DB_ID = "412"
    'Set DB413 = New DBClass    ' stato zona entrata WB (RO)
    'DB413.DB_ID = "413"
  '  Set DB414 = New DBClass    ' stato zona WB (RO)
  '  DB414.DB_ID = "414"
  '  Set DB415 = New DBClass    ' stato zona WB (RO)
  '  DB415.DB_ID = "415"
  '  Set DB416 = New DBClass    ' stato zona smussatura 1 (RO)
  '  DB416.DB_ID = "416"
  '  Set DB417 = New DBClass    ' stato zona smussatura 2 (RO)
  '  DB417.DB_ID = "417"
  '  Set DB418 = New DBClass    ' stato zona cnd (RO)
  '  DB418.DB_ID = "418"
    'Set DB419 = New DBClass    ' stato zona cnd (RO)
    'DB419.DB_ID = "419"
    Set DB420 = New DBClass    ' stato pack pipe (RO)
    DB420.DB_ID = "420"
    Set DB422 = New DBClass    ' stato trasportatori laterali e stoccaggio (RO)
    DB422.DB_ID = "422"
   ' Set DB424 = New DBClass    ' stato zona reggiatura e fasciatura (RO)
   ' DB424.DB_ID = "424"
    Set DB425 = New DBClass    ' stato zona stoccaggio (RO)
    DB425.DB_ID = "425"
    Set DB426 = New DBClass    ' stato zona stoccaggio (RO)
    DB426.DB_ID = "426"
    '===========  comunicazione
    
    Set DB402 = New DBClass    ' comandi vari per comunicazione
    DB402.DB_ID = "402"
    '=========== lettura / scrittura:aree di tracking
    
    Set DB403 = New DBClass    ' commessa futura entrata
    DB403.DB_ID = "403"
    Set DB450 = New DBClass    ' zona caricatore
    DB450.DB_ID = "450"
    Set DB448 = New DBClass    ' zona WB e smussatura
    DB448.DB_ID = "448"
   ' Set DB465 = New DBClass    ' zona CND
   ' DB465.DB_ID = "465"
    Set DB470 = New DBClass    ' dati locali 1mo polmone
    DB470.DB_ID = "470"
'    Set DB471 = New DBClass    ' dati locali 1mo polmone
'    DB471.DB_ID = "471"
    Set DB473 = New DBClass    ' dati locali pale
    DB473.DB_ID = "473"
    'Set DB475 = New DBClass    ' dati locali carrello
    'DB475.DB_ID = "475"
    Set DB480 = New DBClass    ' dati locali reggiatura
    DB480.DB_ID = "480"
    'Set DB481 = New DBClass    ' dati locali reggiatura
    'DB481.DB_ID = "481"
    'Set DB485 = New DBClass    ' dati locali reggiatura
    'DB485.DB_ID = "485"
    Set DB486 = New DBClass    ' dati locali stoccaggio
    DB486.DB_ID = "486"
    
    '============= lettura/scrittura: altri oggetti
    
    Set Cartellino = New LabelClass
    Set ModOrder = New OrderClass
    Set ModRecipe = New RecipeClass
    Set Ricetta = New RecipeClass
  
    '============ Orders e ricette
    
    Set KernelOrder = New OrderClass
    Set KernelRecipe = New RecipeClass
    
    '============ codice ordine
    
    Set CodOrdineCorrente = New COrdineAttuale
    
    '============= recorsets
    
    Set RS_AlarmsLOG = New ADODB.Recordset
    
    ' unità di misura
    Set Unit = New UnitClass
    
 
    '===============================================================================
    'FINE COSTRUZIONE OGGETTI
    '===============================================================================
    
    UploadDatiTurno
    
    'controlla il parametro 210 plc in simulazione
    
    ServerDisattivo = False
    ServerOpcOn = Param.GetBit("Par113_AttivaServerOPC")
    SimulaON = Param.GetBit("Par210_Simulazione_PLCOff")
    
    On Error Resume Next
        
    '============ lettura layout cartellino ====================
    
    Cartellino.Read_file_dati
     
    '******************* lettura iniziale della tabella con i testi degli allarmi
    
    AdoAllarmi.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\AllarmiMultiLingua.mdb;Persist Security Info=False"
    AdoAllarmi.RecordSource = "AllarmiMultiLingua"
    AdoAllarmi.Refresh
    AdoAllarmi.Recordset.ActiveConnection = Nothing

    If Err <> 0 Then MsgBoxA 0, "Dp6 can't read Alarms Database", App.Title, MB_ICONERROR
       
    DoEvents
 
    '===================== LETTURA MAPPA DATI E COSTRUZIONE OGGETTI OPC ================

    Dim IDitem, ItemVar, StringaCollegamento As String
      
'    On Error GoTo ServerError
    If ServerOpcOn Then
       Set OggServer = New OPCServer
       OggServer.Connect (S7_Server)
    End If
    On Error GoTo 0
    
    If Param.GetBit("Par217_AbilitaComLog") Then
       ' DB400.CancLogFile LogComPath
        DB400.LogFile LogComPath, "", True
        DB400.LogErrON = True
        DB410.LogErrON = True
        DB411.LogErrON = True
'        DB412.LogErrON = True
'        DB413.LogErrON = True
      '  DB414.LogErrON = True
      '  DB415.LogErrON = True
      '  DB416.LogErrON = True
      '  DB417.LogErrON = True
      '  DB418.LogErrON = True
'        DB419.LogErrON = True
        DB402.LogErrON = True
        DB403.LogErrON = True
        DB420.LogErrON = True
        DB422.LogErrON = True
  '      DB424.LogErrON = True
        DB425.LogErrON = True
        DB426.LogErrON = True
        DB470.LogErrON = True
        DB450.LogErrON = True
        DB448.LogErrON = True
    '    DB460.LogErrON = True
        DB480.LogErrON = True
        DB486.LogErrON = True
        DB473.LogErrON = True
    Else
        DB400.LogErrON = False
        DB410.LogErrON = False
        DB411.LogErrON = False
        DB448.LogErrON = False
'        DB413.LogErrON = False
    '    DB414.LogErrON = False
    '    DB415.LogErrON = False
    '    DB416.LogErrON = False
    '    DB417.LogErrON = False
    '    DB418.LogErrON = False
'        DB419.LogErrON = False
        DB402.LogErrON = False
        DB403.LogErrON = False
        DB420.LogErrON = False
        DB422.LogErrON = False
  '      DB424.LogErrON = False
        DB425.LogErrON = False
        DB426.LogErrON = False
        DB470.LogErrON = False
        DB450.LogErrON = False
        DB448.LogErrON = False
        DB480.LogErrON = False
        DB486.LogErrON = False
        DB473.LogErrON = False
  '      DB465.LogErrON = False
    End If
    
    ' attivazoni opc e simulazone
    
    DB400.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB402.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB403.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB410.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB411.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
'    DB415.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
'    DB416.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
'    DB417.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
'    DB418.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB420.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB422.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
 '   DB424.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB425.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB426.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB450.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB448.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB470.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB473.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB480.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    DB486.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
'    DB465.Server = Attivo * Abs(ServerOpcOn) + Simula * Abs(SimulaON): ServerDisattivo = SimulaON Or (ServerOpcOn = False)
    '=======================
    Timer1.Enabled = False
    frmPresentazione.PBar1.Max = PB_MAX
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Plc.mdb;Persist Security Info=False"
    cn.Open
    With rs
       .Open "SELECT * FROM Q_ItemCaricati", cn, adOpenKeyset, adLockReadOnly, adCmdText
       .MoveFirst
       While .EOF = False
          IDitem = Left(.Fields("DBItem"), (InStr(.Fields("DBItem"), ",") - 1))
          ItemVar = Right(.Fields("DBItem"), Len(.Fields("DBItem")) - InStr(.Fields("DBItem"), ","))
          StringaCollegamento = S7_nome_collegamento & .Fields("DBItem") & ",1"
          TmpValore = Val(IIf(IsNull(.Fields("Valore")) = False, .Fields("Valore"), 0))
          frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + 1
          frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
          Select Case IDitem
          Case "400"
              DB400.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "402"
              DB402.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "403"
              DB403.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "410"
              DB410.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "411"
              DB411.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "412"
              DB412.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "413"
              DB413.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "414"
              DB414.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "415"
              DB415.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "416"
              DB416.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "417"
              DB417.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "418"
              DB418.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "419"
              DB419.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "420"
              DB420.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "422"
              DB422.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "424"
              DB424.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "425"
              DB425.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "426"
              DB426.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "450"
              DB450.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "448"
              DB448.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "465"
              DB465.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "460"
              DB460.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "470"
              DB470.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "473"
              DB473.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "480"
              DB480.Init IDitem, StringaCollegamento, ItemVar, TmpValore
          Case "486"
              DB486.Init IDitem, StringaCollegamento, ItemVar, TmpValore
         End Select
         .MoveNext
       Wend
       .Close
       Set .ActiveConnection = Nothing
    End With
    
    cn.Close
        
    Set DBTestIN = New DBClass
    Set DBTestOUT = New DBClass
    
    DBTestIN.DB_ID = "DBTestIN"
    DBTestOUT.DB_ID = "DBTestOUT"
    DBTestIN.Server = Attivo * Abs(frmKernel.ServerOpcOn) + Simula * Abs(frmKernel.SimulaON): frmKernel.ServerDisattivo = frmKernel.SimulaON Or (frmKernel.ServerOpcOn = False)
    DBTestOUT.Server = Attivo * Abs(frmKernel.ServerOpcOn) + Simula * Abs(frmKernel.SimulaON): frmKernel.ServerDisattivo = frmKernel.SimulaON Or (frmKernel.ServerOpcOn = False)
    
    Dim StrIO(0 To 1) As String
    Dim k As Integer
    
    ' query di ricerca in mappa dati
    StrIO(0) = "SELECT MappaDati.DBItem, MappaDati.Valore From MappaDati WHERE (((MappaDati.DBItem) Like 'E%') AND ((MappaDati.Gruppo)='Ingressi') AND ((MappaDati.Attivato)=True)) OR (((MappaDati.DBItem) Like 'E%') AND ((MappaDati.Gruppo)='Entrata soffiatura') AND ((MappaDati.Attivato)=True));"
    StrIO(1) = "SELECT MappaDati.DBItem, MappaDati.Valore From MappaDati WHERE (((MappaDati.DBItem) Like 'A%') AND ((MappaDati.Gruppo)='Uscite') AND ((MappaDati.Attivato)=True)) OR (((MappaDati.DBItem) Like 'A%') AND ((MappaDati.Gruppo)='Entrata soffiatura') AND ((MappaDati.Attivato)=True));"
    On Error Resume Next
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Plc.mdb;Persist Security Info=False"
    cn.Open
    For k = 0 To UBound(StrIO)
        With rs
           .Open StrIO(k), cn, adOpenKeyset, adLockReadOnly, adCmdText
           .MoveFirst
           While .EOF = False
              IDitem = "DBTest"
              ItemVar = .Fields("DBItem")
              StringaCollegamento = S7_nome_collegamentoI_O & .Fields("DBItem") & ",1"
              TmpValore = Val(IIf(IsNull(.Fields("Valore")) = False, .Fields("Valore"), 0))
              Select Case Left(ItemVar, 1)
              Case "E"
                  DBTestIN.Init IDitem & "IN", StringaCollegamento, ItemVar, TmpValore
              Case "A"
                  DBTestOUT.Init IDitem & "OUT", StringaCollegamento, ItemVar, TmpValore
             End Select
             .MoveNext
           Wend
           .Close
           Set .ActiveConnection = Nothing
        End With
    Next
    
    Set rs = Nothing
    Set cn = Nothing
    
    
    CaricaPagine
     
    ContaChiusurasys = 120
    ' inizializza l'oggetto grafico mschart
    
    If Not (DB410 Is Nothing) Then InizializzaOggettoGrafico frmEntrata.MSChartTubiMin
    
    'abilitazione temporizzatori
    Timer1.Enabled = True
    TimerHook.Enabled = True
    frmPresentazione.PBar1.value = frmPresentazione.PBar1.Max
    frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    
    Conv_UM.SI_metrico = Param.GetBit("Par101_MisureMetriche")
    
    
    Exit Sub
    
ServerError:
    MsgBox "Please install OPC S7 before !", vbCritical, "Errore d'installazione"
    End
End Sub


Sub ChiusuraProgramma()
   Dim Cancel

   On Error Resume Next
   
' ===================== ABILITAZIONE PASSWORD ALLA CHIUSURA IN COMUNICAZIONE =========
'    If Param.GetBit("Par210_Simulazione_PLCOff") = True Then
'        Cancel = False
'    Else
'        TechPasswordForm.Show (vbModal)
'        If TechPasswordForm.LoginSucceeded Then
'            Cancel = False
'        Else
'            Cancel = True
'        End If
'    End If
'
'    If Cancel = False Then
'==================================================================================
        If ServerOpcOn Then
          OggServer.Disconnect
          Set OggServer = Nothing
        End If
        Set Param = Nothing
        Set Unit = Nothing
        
        Set DB400 = Nothing
        Set DB410 = Nothing
        Set DB411 = Nothing
        Set DB412 = Nothing
        Set DB413 = Nothing
        Set DB414 = Nothing
        Set DB415 = Nothing
        Set DB416 = Nothing
        Set DB417 = Nothing
        Set DB418 = Nothing
        Set DB419 = Nothing
        Set DB420 = Nothing
        Set DB422 = Nothing
        Set DB424 = Nothing
        Set DB425 = Nothing
        Set DB426 = Nothing
        Set DB402 = Nothing
        Set DB403 = Nothing
        Set DB450 = Nothing
        Set DB448 = Nothing
        Set DB465 = Nothing
        Set DB460 = Nothing
        Set DB470 = Nothing
        Set DB471 = Nothing
        Set DB473 = Nothing
        Set DB474 = Nothing
        Set DB475 = Nothing
        Set DB480 = Nothing
        Set DB481 = Nothing
        Set DB486 = Nothing
                
        Set Cartellino = Nothing
        Set ModOrder = Nothing
        Set ModRecipe = Nothing
        Set Ricetta = Nothing
        
        Set KernelOrder = Nothing
        Set KernelRecipe = Nothing
        
        Unload OrderModifyForm
        Set OrderModifyForm = Nothing

'        Unload FormModificaRegge
'        Set FormModificaRegge = Nothing
'
'        Unload RicetteForm
'        Set RicetteForm = Nothing
'
        Unload RecipeModifyForm
        Set RecipeModifyForm = Nothing

        Unload TOUCHKeyBoard
        Set TOUCHKeyBoard = Nothing

        Unload TOUCHNumericPad
        Set TOUCHNumericPad = Nothing

        Unload TechPasswordForm
        Set TechPasswordForm = Nothing

        Unload OrderDeleteForm
        Set OrderDeleteForm = Nothing
        
        End
  '  End If

End Sub

'=================================================================================
'KERNEL cycle
'=================================================================================

'temporizzatore principale del kernel controlla ogni 200ms il programma

Private Sub Timer1_Timer()
Dim LogOK As Boolean
Static oneShot As Boolean
        
        Call Kernel
       
       ' FORZATURA DATACHANGE allo start up dei dati del tracking
       
       If oneShot = False Then
            DB400.Refresh
            DB410.Refresh
            DB411.Refresh
           ' DB414.Refresh
           ' DB415.Refresh
           ' DB416.Refresh
           ' DB417.Refresh
           ' DB418.Refresh
            DB420.Refresh
            DB422.Refresh
      '      DB424.Refresh
            DB425.Refresh
            DB426.Refresh
            DB450.Refresh
       '     DB465.Refresh
       '     DB460.Refresh
            DB473.Refresh
            DB480.Refresh
            'carica i dati delle pagine
            StartupDatiPagina
            'carica le pagine
            Param.ScriviParametriSuPlc
            CommandForm.Show
            CommandForm.PresentazioneOff = False
            'MCCloseForm frmPresentazione, 2
            Unload frmPresentazione
            PaginaCorrente = Pagina.PagLayout
            CommandForm.PresentazioneOff = True
            CommandForm.Aggiornamento
            '============= carica i valori del pacco in corso =========
            LetturaDatiPacco ' aggiorna i dati del cartellino
            '==========================================================
            Set RS_LabelText = New ADODB.Recordset
            frmStampa.LoadTicketFile "..\Target\Tickets\MairTicket_" & Format(Cartellino.Lingua, "00") & ".TKT"
            frmStampa.LoadFixTexts
            frmStampa.LengFixTextRefresh (Cartellino.Lingua)
            frmStampa.FixVisibleRefresh
            '==========================================================
       End If
       
       ' scansione allarmi
       
       Call ScansioneAllarmi
                       
     'ABILITA STORICO DATA CHANGE
                       
     LogOK = Param.GetBit("Par218_LogDataChange")
     
     If Param.GetBit("Par217_AbilitaComLog") Then
        DB400.LogErrON = True
        If LogOK Then DB400.LogDCON = True
        DB402.LogErrON = True
        If LogOK Then DB402.LogDCON = True
        DB403.LogErrON = True
        If LogOK Then DB403.LogDCON = True
        DB410.LogErrON = True
        If LogOK Then DB410.LogDCON = True
        DB411.LogErrON = True
        If LogOK Then DB411.LogDCON = True
      '  DB412.LogErrON = True
      '  If LogOK Then DB412.LogDCON = True
      '  DB413.LogErrON = True
      '  If LogOK Then DB413.LogDCON = True
      '  DB414.LogErrON = True
      '  If LogOK Then DB414.LogDCON = True
      '  DB415.LogErrON = True
      '  If LogOK Then DB415.LogDCON = True
      '  DB416.LogErrON = True
      '  If LogOK Then DB416.LogDCON = True
      '  DB417.LogErrON = True
      '  If LogOK Then DB417.LogDCON = True
      '  DB418.LogErrON = True
      '  If LogOK Then DB418.LogDCON = True
      '  DB419.LogErrON = True
      '  If LogOK Then DB419.LogDCON = True
        DB420.LogErrON = True
        If LogOK Then DB420.LogDCON = True
        DB422.LogErrON = True
        If LogOK Then DB422.LogDCON = True
    '    DB424.LogErrON = True
    '    If LogOK Then DB424.LogDCON = True
        DB425.LogErrON = True
        If LogOK Then DB425.LogDCON = True
        DB426.LogErrON = True
        If LogOK Then DB426.LogDCON = True
        DB450.LogErrON = True
        If LogOK Then DB450.LogDCON = True
        DB448.LogErrON = True
        If LogOK Then DB448.LogDCON = True
      '  DB465.LogErrON = True
      '  If LogOK Then DB465.LogDCON = True
        DB470.LogErrON = True
        If LogOK Then DB470.LogDCON = True
        DB480.LogErrON = True
        If LogOK Then DB480.LogDCON = True
        DB486.LogErrON = True
        If LogOK Then DB486.LogDCON = True
        DB473.LogErrON = True
        If LogOK Then DB473.LogDCON = True
    Else
         DB400.LogErrON = False
         If LogOK Then DB400.LogDCON = False
        DB402.LogErrON = False
        If LogOK Then DB402.LogDCON = False
        DB403.LogErrON = False
        If LogOK Then DB403.LogDCON = False
        DB410.LogErrON = False
        If LogOK Then DB410.LogDCON = False
        DB411.LogErrON = False
        If LogOK Then DB411.LogDCON = False
      '  DB412.LogErrON = False
      '  If LogOK Then DB412.LogDCON = False
      '  DB413.LogErrON = False
      '  If LogOK Then DB413.LogDCON = False
      '  DB414.LogErrON = False
      '  If LogOK Then DB414.LogDCON = False
      '  DB415.LogErrON = False
      '  If LogOK Then DB415.LogDCON = False
      '  DB416.LogErrON = False
      '  If LogOK Then DB416.LogDCON = False
      '  DB417.LogErrON = False
      '  If LogOK Then DB417.LogDCON = False
      '  DB418.LogErrON = False
      '  If LogOK Then DB418.LogDCON = False
      '  DB419.LogErrON = False
      '  If LogOK Then DB419.LogDCON = False
        DB420.LogErrON = False
        If LogOK Then DB420.LogDCON = False
        DB422.LogErrON = False
        If LogOK Then DB422.LogDCON = False
   '     DB424.LogErrON = False
   '     If LogOK Then DB424.LogDCON = False
        DB425.LogErrON = False
        If LogOK Then DB425.LogDCON = False
        DB426.LogErrON = False
        If LogOK Then DB426.LogDCON = False
        DB450.LogErrON = False
        If LogOK Then DB450.LogDCON = False
        DB448.LogErrON = False
        If LogOK Then DB448.LogDCON = False
      '  DB460.LogErrON = False
      '  If LogOK Then DB460.LogDCON = False
        DB470.LogErrON = False
        If LogOK Then DB470.LogDCON = False
        DB480.LogErrON = False
        If LogOK Then DB480.LogDCON = False
        DB486.LogErrON = False
        If LogOK Then DB486.LogDCON = False
        DB473.LogErrON = False
        If LogOK Then DB473.LogDCON = False
    End If
    
    ' assenza di comunicazione
    Static comok As Boolean
    
    If ServerDisattivo = False And DB400.ErroreDB <> "GOOD" And DB402.ErroreDB <> "GOOD" And DB403.ErroreDB <> "GOOD" Then
       If m_PaginaCorrente = PagLayout And comok = False Then frmCOMERROR.Show: comok = True
    End If
    If DB402.Bit(8, 7) Then
       On Error Resume Next
       Shell TargetPath & "Arrestasys.exe", vbNormalFocus
       End
    End If
    oneShot = True
    If SimulaON Then TSimula.Enabled = True
End Sub
'   =================================================================================
'   FUNZIONE DI TRACKING DEI DATI NEL PLC -MAIN KERNEL FUNCTION
'   =================================================================================
Private Sub Kernel()
    Dim i As Integer
    Static oneEndBundle As Boolean
   
    Static CodicePrecDB450 As Integer
    Static CodicePrecDB460 As Integer
    Static CodicePrecDB465 As Integer
    Static CodicePrecDB470 As Integer
    Static CodicePrecDB473 As Integer
    Static CodicePrecDB480 As Integer
    Static CodicePrecDB486 As Integer
    
    Static DB465Valid As Boolean
    Static DB450Valid As Boolean
    Static DB460Valid As Boolean
    Static DB470Valid As Boolean
    Static DB473Valid As Boolean
    Static DB480Valid As Boolean
    Static DB486Valid As Boolean
    
    '================================================================
    ' PRELEVA IL CODICE ATTUALE DAL PLC E LO SCRIVE IN "CODICEPREC"
    '================================================================
    '
    If Not DB465Valid Then
        CodicePrecDB465 = CodOrdineCorrente.CodLav
        DB465Valid = True
    End If

    If Not DB460Valid Then
        CodicePrecDB460 = CodOrdineCorrente.CodWB
        DB460Valid = True
    End If
    If Not DB470Valid Then
        CodicePrecDB470 = CodOrdineCorrente.CodPacco
        DB470Valid = True
    End If
     If Not DB473Valid Then
        CodicePrecDB473 = CodOrdineCorrente.CodMPS
        DB473Valid = True
    End If
    If Not DB480Valid Then
        CodicePrecDB480 = CodOrdineCorrente.CodRegge
        DB480Valid = True
    End If
    If Not DB486Valid Then
        CodicePrecDB486 = CodOrdineCorrente.CodStoccaggio
        DB486Valid = True
    End If
    
    '=================================
    ' scrittura dati su commessa FUTURA in entrata
    '=================================
    
'    Call DownloadOrdiniEvasi
    
    '=================================
    ' scrittura dati su commessa FUTURA in entrata
    '=================================
    
    'KernelOrder.OrdineFuturo è la ricerca del primo codice ordine in DB
    
    KernelOrder.IDOrdine = OrdersForm.CodiceFuturo ' KernelOrder.OrdineFuturo
    
    If KernelOrder.IDOrdine = DB450.Word(0) Then
       PaginaEntrata.Ricetta_Descrizione = KernelOrder.IDRicetta
       PaginaEntrata.Ordine_Descrizione = KernelOrder.Descrizione
       OrdersForm.OrdineTrasmesso KernelOrder.IDOrdine
       DB402.Bit(0, 0) = False
       RefreshListaOrdini = True
    Else
        If KernelOrder.IDOrdine > 0 And DB402.Bit(0, 0) = False Then
            OrdersForm.LoadOrderData KernelOrder
            KernelRecipe.IDRicetta = KernelOrder.IDRicetta
            OrdersForm.LoadRecipeData KernelRecipe
            DownloadDB403 KernelOrder, KernelRecipe
            DB402.Bit(0, 0) = True   ' dati validi
        End If
    End If
    
    If CodicePrecDB450 <> frmKernel.CodOrdineCorrente.CodEntrata Then
       OrderChanged = True
       CodicePrecDB450 = frmKernel.CodOrdineCorrente.CodEntrata
    End If
    
    '================================================================
    ' scrittura dati nuova commessa in WB
    '================================================================
   
   If CodicePrecDB460 <> frmKernel.CodOrdineCorrente.CodWB Then
        KernelOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodWB
        KernelOrder.UploadData frmKernel.CodOrdineCorrente.CodWB
        KernelRecipe.IDRicetta = KernelOrder.IDRicetta
        KernelRecipe.UploadData KernelOrder.IDRicetta
        
        PaginaWb.Ricetta_Descrizione = KernelOrder.IDRicetta
        PaginaWb.Ordine_Descrizione = KernelOrder.Descrizione
        
        If KernelRecipe.UploadData(KernelOrder.IDRicetta) Then
            DownloadDB460 KernelRecipe
            CodicePrecDB460 = frmKernel.CodOrdineCorrente.CodWB
        End If
        OrderChanged = True
    End If
  
  If CodicePrecDB465 <> frmKernel.CodOrdineCorrente.CodLav Then
        KernelOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodLav
        KernelOrder.UploadData frmKernel.CodOrdineCorrente.CodLav
        KernelRecipe.IDRicetta = KernelOrder.IDRicetta
        KernelRecipe.UploadData KernelOrder.IDRicetta
        
        PaginaLav.Ricetta_Descrizione = KernelOrder.IDRicetta
        PaginaLav.Ordine_Descrizione = KernelOrder.Descrizione
        OrderChanged = True
        CodicePrecDB465 = frmKernel.CodOrdineCorrente.CodLav
        'If KernelRecipe.UploadData(KernelOrder.IDRicetta) Then
        '    DownloadDB460 KernelRecipe
        '    CodicePrecDB460 = frmKernel.CodOrdineCorrente.CodWB
        'End If
    End If

    '================================================================
    ' scrittura dati nuova commessa in PACCO : DB470 contiene il codice al cambio ordine
    '================================================================
    '
       
       If CodicePrecDB470 <> frmKernel.CodOrdineCorrente.CodPacco Then
        KernelOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodPacco
        If OrdersForm.LoadOrderData(KernelOrder) Then
            ' carica anche la ricetta del pacco
            KernelRecipe.IDRicetta = KernelOrder.IDRicetta
            OrdersForm.LoadRecipeData KernelRecipe
            ' scrive tutto nel db
            DownloadDB470 KernelOrder, KernelRecipe
            'DownloadDB473 KernelRecipe
            DownloadDB450 KernelRecipe
            CodicePrecDB470 = frmKernel.CodOrdineCorrente.CodPacco
            'OrdersForm.DownloadOrdiniEvasi KernelOrder
            PaginaPacco.Ricetta_Descrizione = KernelOrder.IDRicetta
            PaginaPacco.Ordine_Descrizione = KernelOrder.Descrizione
            OrderChanged = True
        End If
    End If
    
    '================================================================
    ' scrittura dati nuova commessa in PACCO
    '================================================================
    '
    If CodicePrecDB473 <> frmKernel.CodOrdineCorrente.CodMPS Then
        KernelOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodMPS
        If OrdersForm.LoadOrderData(KernelOrder) Then
            KernelRecipe.IDRicetta = KernelOrder.IDRicetta
            OrdersForm.LoadRecipeData KernelRecipe
            DownloadDB473 KernelRecipe
            CodicePrecDB473 = frmKernel.CodOrdineCorrente.CodMPS
            OrderChanged = True
        End If
    End If
    
    '================================================================
    ' scrittura dati nuova commessa in reggiatura
    '================================================================
    '
    ' controlla se nel PLC il codice è cambiato e i dati sono validi
    '
    If CodicePrecDB480 <> frmKernel.CodOrdineCorrente.CodRegge Then
        'carica l'ordine scritto nel plc
        KernelOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodRegge
        '
        If OrdersForm.LoadOrderData(KernelOrder) Then
             ' se sono stati caricati i dati dell'ordine allora carica anche i dati della ricetta
             ' l'indice ricetta presente nell'ordine viene copiato nell'id ricetta (relazione)
            KernelRecipe.IDRicetta = KernelOrder.IDRicetta
            ' carica i dati della ricetta
            OrdersForm.LoadRecipeData KernelRecipe
            ' scarica i dati della ricetta nel db480
            DownloadD_ricetta_B480 KernelRecipe
            'aggiorna il codice precedente
            CodicePrecDB480 = frmKernel.CodOrdineCorrente.CodRegge
            PaginaReggiatura.Ricetta_Descrizione = KernelOrder.IDRicetta
            PaginaReggiatura.Ordine_Descrizione = KernelOrder.Descrizione
        End If
        OrderChanged = True
    End If

    '================================================================
    ' scrittura dati nuova commessa in stoccaggio
    '================================================================
    '
    
    If CodicePrecDB486 <> frmKernel.CodOrdineCorrente.CodStoccaggio Then
        KernelOrder.IDOrdine = frmKernel.CodOrdineCorrente.CodStoccaggio
        If OrdersForm.LoadOrderData(KernelOrder) Then
            KernelRecipe.IDRicetta = KernelOrder.IDRicetta
            OrdersForm.LoadRecipeData KernelRecipe
'            If Param.GetBit("Par229_STORAGE_DEST") Then
'               DB426.Word(70) = KernelRecipe.Destination
'            End If
            CodicePrecDB486 = frmKernel.CodOrdineCorrente.CodStoccaggio
            PaginaStoccaggio.Ricetta_Descrizione = KernelOrder.IDRicetta
            PaginaStoccaggio.Ordine_Descrizione = KernelOrder.Descrizione
        End If
        OrderChanged = True
    End If
    
    '=================================================================
    ' archiviazione pacco
    '=================================================================
    
    'ristampa del cartellino da pulsante esterno
    If DB402.Bit(8, 5) Then
       If Param.GetBit("Par112_StampaInserita") And PrinterInstall Then
            With frmStampa
               ' .LoadTicketFile "..\Target\Tickets\MairTicket_" & Format(Cartellino.Lingua, "00") & ".TKT"
''                .LoadFixTexts
''                .LengFixTextRefresh (Cartellino.Lingua)
''                .FixVisibleRefresh
''                .RefreshVar
                .PrintExec DB402.Word(18)
                RecipeModifyForm.TicketFrefresh = False
            End With
        End If
        DB402.Bit(8, 1) = DB402.Bit(8, 4)
        DB402.Bit(8, 5) = False
    End If
    
    BitPesa = ((DB402.Bit(8, 0) = True) And (DB402.Bit(8, 1) = False) Or DB402.Bit(8, 4))
    
    ' salvataggio dati pacco : DB402.Bit(8, 2) peso valido dal plc
    Static relay As Integer
    
    
    If BitPesa Or Param.GetBit("Par219_ArchivForzata") Then       ' richiesta archiviazione da PLC
        ' aggiorna campo automatici dell'oggetto "Cartellino" e salva su storico
        If relay < 1 Then
        frmKernel.PaccoArchiviato = False
        DB402.Bit(8, 1) = True  ' set archiviazione eseguita a PLC
        If Param.GetBit("Par219_ArchivForzata") Then Param.SetBit "Par219_ArchivForzata", False
        'frmStampa.Aggiorna = True
        Call ArchiviaPacco
'        Call DownloadOrdiniEvasi
        PaccoArchiviato = True
        If Param.GetBit("Par112_StampaInserita") And PrinterInstall Then
           'For i = 1 To Param.GetNumber("Par221_NumeroCartellini")
            With frmStampa
               ' .LoadTicketFile "..\Target\Tickets\MairTicket_" & Format(Cartellino.Lingua, "00") & ".TKT"
''                .LoadFixTexts
''                .LengFixTextRefresh (Cartellino.Lingua)
''                .FixVisibleRefresh
''                .RefreshVar
                .PrintExec DB402.Word(18)
                RecipeModifyForm.TicketFrefresh = False
            End With
           'Next
        End If
        DoEvents
        Dim localdate As String
        localdate = Format(Month(Now), "00") & "/" & Format(Day(Now), "00") & "/" & Format(year(Now), "0000")
        BundlesLogForm.Filtro = "SELECT * FROM Bundles WHERE ((DataOra >= '" & localdate & _
                   " 00.00.00' ) AND (DataOra <= '" & localdate & " 23.59.59' )) ORDER BY Contatore;"
        BundlesLogForm.Aggiorna
        relay = 1
        End If
        
   End If
    
   If relay > 0 Then
      relay = relay + 1
   End If
   If relay > 10 Then relay = 0
    
   DB402.Refresh
   
   If DB402.Bit(0, 3) And Not oneEndBundle Then
        frmKernel.PaginaCorrente = PagPacco
        oneEndBundle = True
   End If
   If Not DB402.Bit(0, 3) Then
        oneEndBundle = False
   End If

End Sub


'**************************************************
' Funzione di aggiornamento video e comunicazione
' con il plc da chiamare in background
'**************************************************

'**************************************************
' Funzione di aggiornamento video e comunicazione
' con il plc da chiamare in background
'**************************************************
Public Sub ScansioneAllarmi()
  Static jump As Boolean
  Dim Salvataggio As Boolean
  
  jump = Not jump
  If jump = False Then Exit Sub
  
  Salvataggio = False
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB400", DB400, EcoAllarmiDB400, ListaAllarmiDB400, ListaMessaggiDB400)
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB410", DB410, EcoAllarmiDB410, ListaAllarmiDB410, ListaMessaggiDB410)
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB411", DB411, EcoAllarmiDB411, ListaAllarmiDB411, ListaMessaggiDB411)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi "DB412", DB412, EcoAllarmiDB412, ListaAllarmiDB412, ListaMessaggiDB412
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB413", DB413, EcoAllarmiDB413, ListaAllarmiDB413, ListaMessaggiDB413)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB414", DB414, EcoAllarmiDB414, ListaAllarmiDB414, ListaMessaggiDB414)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB415", DB415, EcoAllarmiDB415, ListaAllarmiDB415, ListaMessaggiDB415)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB416", DB416, EcoAllarmiDB416, ListaAllarmiDB416, ListaMessaggiDB416)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB417", DB417, EcoAllarmiDB417, ListaAllarmiDB417, ListaMessaggiDB417)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB418", DB418, EcoAllarmiDB418, ListaAllarmiDB418, ListaMessaggiDB418)
'  Salvataggio = Salvataggio Or ScansioneDBAllarmi "DB419", DB419, EcoAllarmiDB419, ListaAllarmiDB419, ListaMessaggiDB419
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB420", DB420, EcoAllarmiDB420, ListaAllarmiDB420, ListaMessaggiDB420)
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB422", DB422, EcoAllarmiDB422, ListaAllarmiDB422, ListaMessaggiDB422)
'   Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB424", DB424, EcoAllarmiDB424, ListaAllarmiDB424, ListaMessaggiDB424)
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB425", DB425, EcoAllarmiDB425, ListaAllarmiDB425, ListaMessaggiDB425)
  Salvataggio = Salvataggio Or ScansioneDBAllarmi("DB426", DB426, EcoAllarmiDB426, ListaAllarmiDB426, ListaMessaggiDB426)

  If m_Allarme = False Then AllarmeON = False
  Exit Sub
Aperto:


End Sub

Private Function ScansioneDBAllarmi(DBName As String, DataBlock As DBClass, DataBlockBuffer() As Boolean, ListaAllarmi() As String, ListaMessaggi() As String) As Boolean
    Dim ByteInizioAllarmi As Integer
    Dim ByteInizioMessaggi As Integer
    Dim ByteFinale As Integer
    
    Dim ByteCounter As Integer
    Dim BitCounter As Integer
    Dim AlarmWord As Integer
    Dim BufferWord As Integer
    Dim IndiceAllarme As Integer
    Dim IndiceMessaggio As Integer
    Dim MsByte As Integer
    Dim Testo As String
    Dim Id_Item As Integer
    Dim i, j As Integer
    
    'On Error GoTo Fine
    
    ' Azzeramento lista allarmi attivi
    For IndiceAllarme = 0 To MaxAllarmi
        ListaAllarmi(IndiceAllarme) = ""
        ListaMessaggi(IndiceAllarme) = ""
    Next IndiceAllarme
    ' scansione degli allarmi bit per bit
    
    IndiceAllarme = 0
    IndiceMessaggio = 0
    ScansioneDBAllarmi = False
    
    'area allarmi e messaggi
    ByteInizioAllarmi = 1
    ByteFinale = 24
    ByteInizioMessaggi = 18
    
    Dim pvByte As Integer
    Dim pvBit As Integer
    Dim StatoBit As Boolean
    
    For i = 2 To 16 Step 2
      For j = 0 To 15
         If DataBlock.MaskBit(i, j, True) <> 2 Then
                pvByte = i + Int(j / 8)
                pvBit = j - 8 * Int(j / 8)
                StatoBit = DataBlock.MaskBit(i, j)
                If StatoBit Then
                   ListaAllarmi(IndiceAllarme) = TestoAllarme(DBName, pvByte, pvBit)
                   If Not DataBlockBuffer(pvByte, pvBit) Then 'controlla se lo stato precedente era basso o alto
                       LogAlarm Left(ListaAllarmi(IndiceAllarme), InStr(ListaAllarmi(IndiceAllarme), ":") - 1), True ' lo salva nel batabase (non registra i messaggi)
                       NumAllAttivi = NumAllAttivi + 1
                       ScansioneDBAllarmi = True
                   End If
                   IndiceAllarme = IndiceAllarme + 1
                   AllarmeON = True
                   m_Allarme = True
                Else
                   If DataBlockBuffer(pvByte, pvBit) Then
                       ListaAllarmi(IndiceAllarme) = TestoAllarme(DBName, pvByte, pvBit)
                       LogAlarm Left(ListaAllarmi(IndiceAllarme), InStr(ListaAllarmi(IndiceAllarme), ":") - 1), False
                       m_Allarme = False
                       ListaAllarmi(IndiceAllarme) = ""
                       NumAllAttivi = NumAllAttivi - 1
                       ScansioneDBAllarmi = True
                   End If
                End If
                DataBlockBuffer(pvByte, pvBit) = StatoBit
         End If
      Next
    Next
    ' messaggi
   For i = 18 To 24 Step 2
      For j = 0 To 15
         If DataBlock.MaskBit(i, j, True) <> 2 Then
                pvByte = i + Int(j / 8)
                pvBit = j - 8 * Int(j / 8)
                StatoBit = DataBlock.MaskBit(i, j)
                If StatoBit Then
                    ListaMessaggi(IndiceMessaggio) = TestoAllarme(DBName, pvByte, pvBit)
                    If Not DataBlockBuffer(pvByte, pvBit) Then  'controlla se lo stato precedente era basso o alto
                        NumMsgAttivi = NumMsgAttivi + 1
                    End If
                    AllarmeON = True
                    m_Allarme = True
                    IndiceMessaggio = IndiceMessaggio + 1
                Else
                    If DataBlockBuffer(pvByte, pvBit) Then
                       ListaMessaggi(IndiceMessaggio) = ""
                       NumMsgAttivi = NumMsgAttivi - 1
                       m_Allarme = False
                    End If
                End If
                DataBlockBuffer(pvByte, pvBit) = StatoBit
         End If
      Next
    Next

Fine:
    DataBlock.DatiCambiati = False
    DoEvents
End Function

Private Function TestoAllarme(DBName As String, IndiceByte As Integer, IndiceBit As Integer) As String
    Dim AlarmID As String
    Dim CampoLingua As String
    
    Select Case Param.GetNumber("Par100_Lingua")
        Case 1
            CampoLingua = "ITALIANO"
        Case 2
            CampoLingua = "INGLESE"
        Case 3
            CampoLingua = "FRANCESE"
        Case 4
            CampoLingua = "SPAGNOLO"
        Case 5
            CampoLingua = "TEDESCO"
        Case 6
            CampoLingua = "LinguaSpeciale"
        Case Else
            CampoLingua = "ITALIANO"
    End Select
    AlarmID = DBName & "_" & Format(IndiceByte, "000") & "_" & IndiceBit
    AdoAllarmi.Recordset.MoveFirst
    AdoAllarmi.Recordset.Find ("TagName = '" & AlarmID & "'")
    If AdoAllarmi.Recordset.EOF = False Then
       If (AdoAllarmi.Recordset.Fields("Riferimento") = "") Or IsNull(AdoAllarmi.Recordset.Fields("Riferimento")) Then
           TestoAllarme = AlarmID & " : " & AdoAllarmi.Recordset.Fields(CampoLingua)
       Else
           TestoAllarme = AlarmID & " : " & AdoAllarmi.Recordset.Fields(CampoLingua) & " [" & AdoAllarmi.Recordset.Fields("Riferimento") & "]"
       End If
    Else
        TestoAllarme = AlarmID & " : "
    End If
End Function
Public Function HelpAllarme(DBName As String, IndiceByte As Integer, IndiceBit As Integer) As String
    Dim AlarmID As String
    Dim CampoLingua As String
    Dim AlFile As String
    
    Select Case Param.GetNumber("Par100_Lingua")
        Case 1
            CampoLingua = "ITALIANO"
        Case 2
            CampoLingua = "INGLESE"
        Case 3
            CampoLingua = "FRANCESE"
        Case 4
            CampoLingua = "SPAGNOLO"
        Case 5
            CampoLingua = "TEDESCO"
        Case 6
            CampoLingua = "LinguaSpeciale"
        Case Else
            CampoLingua = "ITALIANO"
    End Select
    AlarmID = DBName & "_" & Format(IndiceByte, "000") & "_" & IndiceBit
    AdoAllarmi.Recordset.MoveFirst
    AdoAllarmi.Recordset.Find ("TagName = '" & AlarmID & "'")
    ' aggiungere campolingua per restituire allarmi multilingua
    
    HelpAllarme = "NULL"
    If AdoAllarmi.Recordset.EOF = False Then
       AlFile = AdoAllarmi.Recordset.Fields("HelpFile") & ""
       If AlFile <> "" Then HelpAllarme = AlFile
    End If
    If HelpAllarme <> "NULL" Then HelpAllarme = HelpAllarme & ".htm"
End Function
Private Sub LogAlarm(Testo As String, Active As Boolean)
'    Dim i As Integer
'    Dim MaxAlarmLog As Long
'    Dim Trovato As Boolean
'
'    On Error Resume Next
'
'    With RS_AlarmsLOG
'        If FileEsistente("..\target\AlarmsLOG.xml") Then
'           If .State = adStateClosed Then
'              .Open "..\target\AlarmsLOG.xml", , , , adCmdFile
'              If Err <> 0 Then
'                 Kill "..\target\AlarmsLOG.xml"
'                 If .State = adStateOpen Then .Close
'                 Open "..\target\LogErrori.txt" For Append As #2
'                 Print #2, Format(Now, "dd-mm-yyyy hh:mm:ss") & "Repair AlarmsLOG (kill)"
'                 Close #2
'                 Exit Sub
'              End If
'           End If
'        Else
'           If .State = adStateOpen Then .Close
'           .Fields.Append "Data", adChar, 20
'           .Fields.Append "Ora", adChar, 20
'           .Fields.Append "Turno", adChar, 200
'           .Fields.Append "Stato", adChar, 5
'           .Fields.Append "Descrizione", adChar, 255
'           .Open
'           .Save "..\target\AlarmsLOG.xml", adPersistXML
'           .Close
'        End If
'        ' controllo recordset aperto
'        DoEvents
'        If .State = adStateClosed Then
'           .Open "..\target\AlarmsLOG.xml", , , , adCmdFile
'        End If
'        '======================================
'        '     cancellazione 90 record vecchi
'        MaxAlarmLog = Param.GetNumber("Par203_AllarmiSuStorico")
'        If MaxAlarmLog < 100 Then MaxAlarmLog = 100 'numero minimo di allarmi
'        If .RecordCount > MaxAlarmLog Then
'           For i = 0 To 90
'               .MoveFirst
'               .Delete
'           Next i
'        End If
'        ' registrazione nuovo record
'        '======================================
'        .AddNew
'        .Fields("Data") = Month(Now) & "/" & Day(Now) & "/" & year(Now)
'        .Fields("Ora") = Hour(Now) & "." & Minute(Now) & "." & Second(Now)
'        ' scrive il turno
'        Trovato = False
'        For i = 1 To 3
'           If Hour(Time) >= Val(DatiTurno(i, 1)) And Hour(Time) < Val(DatiTurno(i, 2)) Then
'              .Fields("Turno") = UCase(DatiTurno(i, 3))
'              Trovato = True
'              Exit For
'           End If
'        Next
'        If Trovato = False Then
'           For i = 1 To 3
'              If (Hour(Time) > Val(DatiTurno(i, 1)) And Hour(Time) <= 23) Or (Hour(Time) >= 0 And Hour(Time) < Val(DatiTurno(i, 2))) Then
'                 .Fields("Turno") = UCase(DatiTurno(i, 3))
'                 Trovato = True
'                 Exit For
'              End If
'           Next
'        End If
'        If Trovato = False Then
'           .Fields("Turno") = "NULL"
'        End If
'
'        If Active Then
'            .Fields("Stato") = "ON"
'        Else
'            .Fields("Stato") = "OFF"
'        End If
'        .Fields("Descrizione") = Testo
'        .Update
'        '======================================
'    End With
    '============================================================
    Dim i As Integer
    Dim MaxAlarmLog As Long
    On Error Resume Next
    MaxAlarmLog = Param.GetNumber("Par203_AllarmiSuStorico")
    AlarmLog.removeElementiFile TargetPath & "Alarmslog.txt", MaxAlarmLog
    AlarmLog.addNewErr Testo, Active, Val(DatiTurno(1, 1)), _
                       Val(DatiTurno(1, 2)), Val(DatiTurno(2, 1)), _
                       Val(DatiTurno(2, 2)), Val(DatiTurno(3, 1)), _
                       Val(DatiTurno(3, 2))
    AlarmLog.Scrivi TargetPath & "Alarmslog.txt"
    '============================================================
End Sub

Sub ArchiviaPacco()
    Dim Ordine As OrderClass
    Dim tempRicetta As New RecipeClass
    Dim i As Integer
    Dim MaxBundleLog As Long
    Dim FlagComapattazione As Boolean
    Dim bundleNumber As String
    
    On Error Resume Next
    If RipristinaDatabase = False Then
        If Dir("..\target\HistoryCopy.mdb") <> "" Then
           Kill "..\target\HistoryCopy.mdb"
        End If
        FileCopy "..\target\History.mdb", "..\target\HistoryCopy.mdb"
    End If
    
    AdoLogPacco.Refresh
        
    ' cancellazione 10 record vecchi
    MaxBundleLog = Param.GetNumber("Par202_PacchiSuStorico")
    If MaxBundleLog < 20 Then MaxBundleLog = 20
    If AdoLogPacco.Recordset.RecordCount > MaxBundleLog Then
        For i = 0 To 10
           AdoLogPacco.Recordset.MoveFirst
           AdoLogPacco.Recordset.Delete
        Next i
        ' memoria compattazione
        FlagComapattazione = True
    End If
    
    '******************************* registrazione nuovo record ************************************
    
    AdoLogPacco.Recordset.AddNew
    If Err <> 0 Then
       AdoLogPacco.Recordset.ActiveConnection = Nothing
       Exit Sub
    End If
    
    '===========================================================================
    '                                              AGGIORNA I CAMPI DEL CARTELLINO
    '===========================================================================
    ' legge ordine da plc
    Set Ordine = New OrderClass
    Ordine.IDOrdine = DB402.Word(18)   'il codice dello stoccaggio si trova nel 402
    OrdersForm.LoadOrderData Ordine
    tempRicetta.UploadData Ordine.IDRicetta
    PaginaArchivia.Ricetta_Descrizione = Ordine.IDRicetta
    PaginaArchivia.Ordine_Descrizione = Ordine.Descrizione
    '===========================================================================
    ' trasferimento sul cartellino dei dati dell'ordine
    '===========================================================================
    
    Cartellino.Descrizione = Ordine.Descrizione
    ' aggiorna campo data e ora
    Cartellino.Data = Format(Now, "mm/dd/yyyy")
    Cartellino.Ora = Format(Now, "hh:mm")
    ' aggiorna altri campi automatici con dati da plc
    Cartellino.NumeroPacco = Format(DB402.Word(12), "###0")
    Cartellino.NumeroTubi = Format(DB402.Word(14), "###0")
    Cartellino.DimensioniTubo = Ordine.IDRicetta  ' Unit.m_To_Display_mm0(DB402.Word(24) / 10000#) & " X " & Unit.m_To_Display_mm0(DB402.Word(26) / 10000#)   ' altezza x larghezza  (decimi)
    Cartellino.LunghezzaTubo = Conv_UM.Conversione(tempRicetta.TuboLunghezza, UM.mt, UM.ft, 0) ' Unit.m_To_Display_mm(DB402.Word(22) / 1000#)
    Cartellino.SpessoreTubo = Conv_UM.Conversione(tempRicetta.TuboSpessore, UM.mt, UM.inch, 4) 'Unit.m_To_Display_mm00((DB402.Word(28) / 100000#)) * 10#  ' & Unit.mmString
'    Cartellino.PesoPacco = Conv_UM.Conversione(tempRicetta.PaccoPesoTeorico, UM.kg, UM.lB, 3)
    Cartellino.PesoPacco = Round(Conv_UM.Conversione(tempRicetta.TuboLunghezza, UM.mt, UM.ft, 0) * tempRicetta.WeightPerFeet * Cartellino.NumeroTubi, 0)
    WeightToPrint = Cartellino.PesoPacco
    'peso del pacco
'    If DB402.Word(16) >= 0 And Param.GetBit("Par204_AttivaPesa") Then
'       Cartellino.PesoPacco = Conv_UM.Conversione(DB402.Word(16), UM.kg, UM.lB)
'    Else
'      Cartellino.PesoPacco = "---"
'    End If
    
    '===========================================================================
    ' composizione cartellino
    '===========================================================================
    With Cartellino
        .Read_file_dati 'carica i dati dal file
        For i = 1 To 10
           Cartellino.CampoManuale(i) = Ordine.CampoManuale(i)
        Next i
        .CampoAuto(1) = .Data
        .CampoAuto(2) = .Ora
        .CampoAuto(3) = .NumeroPacco
        .CampoAuto(4) = .NumeroTubi
        .CampoAuto(5) = .DimensioniTubo
        .CampoAuto(6) = .LunghezzaTubo
        .CampoAuto(7) = .SpessoreTubo
        .CampoAuto(8) = .PesoPacco
        .CampoAuto(9) = .Descrizione
        .Scrive_file_dati 'salva i dati nel file
    End With
    '===========================================================================
    ' export dati archiviazione su file di testo
    '===========================================================================
    Dim outstring As String
    
    bundleNumber = Export.UniqueKeyAdd(" 5")
    
    Export.Increase_Exports_number
    
    'esportazione storico pacchi
    
    outstring = ""
    outstring = outstring & Chr(34) & Cartellino.Data & " " & Cartellino.Ora & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.Data & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.Ora & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.Descrizione & Chr(34) & ","
    For i = 1 To 10
        outstring = outstring & Chr(34) & Cartellino.CampoManuale(i) & Chr(34) & ","
    Next i
    outstring = outstring & Chr(34) & Cartellino.NumeroPacco & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.NumeroTubi & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.DimensioniTubo & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.LunghezzaTubo & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.SpessoreTubo & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.PesoPacco & Chr(34) & ","
    outstring = outstring & Chr(34) & DB402.DWord(40) & Chr(34) & "#"
    
    Export.EXPORT_BundleLog outstring
    
    'composizione stringa uscita
    Dim tmpStr As String
    
    If Trim(Ordine.CampoManuale(4)) <> "" Then
        tmpStr = Ordine.CampoManuale(3) & "," & Ordine.CampoManuale(4)
    Else
        tmpStr = Ordine.CampoManuale(3)
    End If
    
    outstring = ""
    outstring = outstring & bundleNumber & ","
    outstring = outstring & Chr(34) & tmpStr & Chr(34) & ","
'    outstring = outstring & Chr(34) & Ordine.CampoManuale(4) & Chr(34) & ","
    outstring = outstring & Chr(34) & Ordine.CampoManuale(8) & Chr(34) & ","
    outstring = outstring & Chr(34) & Ordine.CampoManuale(2) & Chr(34) & ","
    outstring = outstring & Chr(34) & Cartellino.CampoAuto(6) & Chr(34) & ","
    outstring = outstring & Chr(34) & Chr(34) & ","
    outstring = outstring & Chr(34) & Ordine.CampoManuale(7) & Chr(34) & ","
    outstring = outstring & Chr(34) & Ordine.CampoManuale(5) & Chr(34) & ","
    outstring = outstring & Chr(34) & Ordine.CampoManuale(1) & Chr(34) & ","
    outstring = outstring & "#" & year(Now) & "-" & Day(Now) & "-" & Month(Now) & "#"
    
    'appende la stringa nel file
    
    Export.SaveBundleData outstring
    
    '===========================================================================
    ' salva su database STORICO il cartellino
    '===========================================================================
    
'    Ordine.CampoManuale(4) = ModRecipe.IDRicetta
'''    Ricetta.IDRicetta = ModRecipe.IDRicetta
'    Ordine.CampoManuale(2) = ModRecipe.Grade
''    Ricetta.Grade = ModRecipe.Grade
'    Ordine.CampoManuale(5) = ModRecipe.Pieces
''    Ricetta.Pieces = ModRecipe.Pieces
'    Ordine.CampoManuale(8) = ModRecipe.Itemcode
''    Ricetta.Itemcode = ModRecipe.Itemcode
'    Ordine.CampoManuale(7) = ModRecipe.PaccoPesoTeorico
    
    
    AdoLogPacco.Recordset.Fields("DataOra") = Cartellino.Data & " " & Cartellino.Ora
    AdoLogPacco.Recordset.Fields("Data") = Cartellino.Data
    AdoLogPacco.Recordset.Fields("Ora") = Cartellino.Ora
    AdoLogPacco.Recordset.Fields("OrdineDescrizione") = Cartellino.Descrizione
    For i = 1 To 10
        AdoLogPacco.Recordset.Fields("Cartellino" & i) = Cartellino.CampoManuale(i)
    Next i
    AdoLogPacco.Recordset.Fields("NumeroPacco") = Cartellino.NumeroPacco
    AdoLogPacco.Recordset.Fields("NumeroTubi") = Cartellino.NumeroTubi
    AdoLogPacco.Recordset.Fields("DimensioniTubo") = Cartellino.DimensioniTubo
    AdoLogPacco.Recordset.Fields("LunghezzaTubo") = Cartellino.LunghezzaTubo
    AdoLogPacco.Recordset.Fields("SpessoreTubo") = Cartellino.SpessoreTubo
    AdoLogPacco.Recordset.Fields("PesoPacco") = Cartellino.PesoPacco
    AdoLogPacco.Recordset.Fields("Cartellino9") = DB402.DWord(40)
    
    AdoLogPacco.Recordset.Update
    AdoLogPacco.Recordset.ActiveConnection = Nothing
    
    If FlagComapattazione Then
        CompattazioneDB "..\target\History.mdb"
    End If
    
End Sub

Private Sub TimerHook_Timer()
   Static Minuto As Variant
   Static FirstScan As Boolean
   Dim NewValue As Long
   Dim i As Long
   
   'controllo contatore tubi in entrata

   If (Not (DB410 Is Nothing)) Then ' se è caricato il db410
      If Minuto <> Minute(Time) Then ' cambia ogni minuto
         If FirstScan = False Then
            DB410.Word(32) = 0
            FirstScan = True
            NewValue = 0
         Else
         ' aggiorna i minuti
         BufferTubiMin.MinuteActual = Minute(Time)
         BufferTubiMin.Data = Format(Date$, "mm-dd-yyyy") & "   " & Format(Time$, "hh.mm")
         ' shift dati precedenti
         For i = (BufferTubiMin.Minutelog - 1) To 1 Step -1
              BufferTubiMin.BufferData(i + 1, 1) = BufferTubiMin.BufferData(i, 1)
         Next
         ' inserimento nuovo dato letto
         NewValue = DB410.Word(32)
         DB410.Word(32) = 0
         BufferTubiMin.BufferData(1, 1) = NewValue
         Minuto = Minute(Time)
         End If
         ' aggiorna il grafico
         DownloadDatiGrafico frmEntrata.MSChartTubiMin
      End If
   
   End If
   
   'Visualizza stato,ora e data correnti
    CommandForm.LblData = Format(Day(Date), "00") & "/" & Format(Month(Date), "00") & "/" & year(Date)
    CommandForm.LblOra = Time$

   ' controlla se è premuta la combinazione di tasti Ctrl+Alt+E
    If GetAsyncKeyState(vbKeyE) And &H8000 Then
        If GetAsyncKeyState(vbKeyControl) And &H8000 Then
            If GetAsyncKeyState(vbKeyMenu) And &H8000 Then
                ' se la combinazione è OK lancia un fine programma
                If PaginaCorrente = PagLayout Then Exit Sub
                DoEvents
                On Error Resume Next
                Unload frmCOMERROR
                On Error GoTo Errore
                TechPasswordForm.defPassWord = Trim(Param.GetNumber("Par111_Password utente"))
                TechPasswordForm.Show vbModal
                If TechPasswordForm.LoginSucceeded = False Then Exit Sub
                Unload TechPasswordForm
                If Param.GetBit("Par211_DownloadPLCDataInDB") Then
                    frmMovingData.FromPLC = True
                    frmMovingData.Show
                    Call SalvaDatiPLCInDBSimulazione
                    frmMovingData.Hide
                 End If
                 Call ChiusuraProgramma
                 End
            End If
        End If
    End If
   
    'aggiorna lo stato del pulsante allarmi della barra comandi
    
    On Error Resume Next
    
    If AllarmeON = True Then
        frmStatistica.Barra21.Allarme = True
        frmEntrata.Barra21.Allarme = True
      '  FormLavaggio.Barra21.Allarme = True
        AlarmForm.Barra21.Allarme = True
      '  FilettoForm.Barra21.Allarme = True
        CommandForm.Barra21.Allarme = True
        OrdersForm.Barra21.Allarme = True
        BundleForm.Barra21.Allarme = True
        frmStoccaggio.Barra21.Allarme = True
        FormRegge.Barra21.Allarme = True
        Param.Barra21.Allarme = True
      '  frmSmussatrice.Barra21.Allarme = True
        BundlesLogForm.Barra21.Allarme = True
      '  FormWB.Barra21.Allarme = True
    Else
        frmStatistica.Barra21.Allarme = False
        frmEntrata.Barra21.Allarme = False
        AlarmForm.Barra21.Allarme = False
      '  FormLavaggio.Barra21.Allarme = False
      '  FilettoForm.Barra21.Allarme = False
        CommandForm.Barra21.Allarme = False
        OrdersForm.Barra21.Allarme = False
        BundleForm.Barra21.Allarme = False
        frmStoccaggio.Barra21.Allarme = False
        FormRegge.Barra21.Allarme = False
        Param.Barra21.Allarme = False
    '    frmSmussatrice.Barra21.Allarme = False
        BundlesLogForm.Barra21.Allarme = False
     '   FormWB.Barra21.Allarme = False
    End If
Errore:
    On Error GoTo 0
End Sub

' Rende visibile lo stato del ciclo timer

Property Get StatoCom() As Boolean
Dim ErroriDB As Boolean
Dim ErroriCom As Boolean

        ErroriDB = False
        ErroriCom = False
        
        'errori nei db
        ErroriDB = DB400.ErroreDB <> "GOOD" Or DB402.ErroreDB <> "GOOD" Or DB403.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB411.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB420.ErroreDB <> "GOOD" Or DB422.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB470.ErroreDB <> "GOOD" Or DB486.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB473.ErroreDB <> "GOOD" Or DB480.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB450.ErroreDB <> "GOOD" Or DB410.ErroreDB <> "GOOD"
        'ErroriDB = ErroriDB Or DB460.ErroreDB <> "GOOD" Or DB416.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB425.ErroreDB <> "GOOD" Or DB448.ErroreDB <> "GOOD"
        ErroriDB = ErroriDB Or DB426.ErroreDB <> "GOOD"
        'ErroriDB = ErroriDB Or DB465.ErroreDB <> "GOOD"
        
        'errori comunicazione
        ErroriCom = DB400.ComError <> "" Or DB402.ComError <> "" Or DB403.ComError <> ""
        ErroriCom = ErroriCom Or DB422.ComError <> "" Or DB470.ComError <> "" Or DB420.ComError <> ""
        ErroriCom = ErroriCom Or DB473.ComError <> "" Or DB480.ComError <> "" Or DB486.ComError <> ""
        ErroriCom = ErroriCom Or DB450.ComError <> "" Or DB410.ComError <> "" 'Or DB415.ComError <> ""
        'ErroriCom = ErroriCom Or DB416.ComError <> "" Or DB417.ComError <> "" Or DB410.ComError <> ""
        ErroriCom = ErroriCom Or DB425.ComError <> "" Or DB448.ComError <> ""
        ErroriCom = ErroriCom Or DB426.ComError <> ""
        ' stato del semaforo di comunicazione
        StatoCom = ErroriCom Or ErroriDB
End Property
Sub SalvaDatiPLCInDBSimulazione()
        Dim IDitem, ItemVar As String
        Dim valore As Variant
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        
        On Error GoTo Passo2
        
Passo1: If FileEsistente("..\target\Simulazioneplc.mdb") Then Kill "..\target\Simulazioneplc.mdb"
        While FileEsistente("..\target\Simulazioneplc.mdb")
           DoEvents
        Wend
        
        FileCopy "..\target\plc.mdb", "..\target\Simulazioneplc.mdb"
        While FileEsistente("..\target\Simulazioneplc.mdb") = False
           DoEvents
        Wend
        
Passo2: On Error GoTo Errore
        cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Plc.mdb;Persist Security Info=False"
        cn.Open
        With rs
               .Open "SELECT * FROM MappaDati", cn, adOpenKeyset, adLockOptimistic, adCmdText
               .MoveFirst
               While .EOF = False
                  DoEvents
                  If .Fields("Attivato") Then
                         IDitem = Left(.Fields("DBItem"), 3)
                         ItemVar = Right(.Fields("DBItem"), Len(.Fields("DBItem")) - 4)
                         Select Case IDitem
                         Case "400"
                              valore = DB400.Item(ItemVar)
                         Case "402"
                              valore = DB402.Item(ItemVar)
                         Case "403"
                              valore = DB403.Item(ItemVar)
                         Case "410"
                              valore = DB410.Item(ItemVar)
                         Case "411"
                              valore = DB411.Item(ItemVar)
                         Case "412"
                              valore = DB412.Item(ItemVar)
                         Case "413"
                              valore = DB413.Item(ItemVar)
                         Case "414"
                              valore = DB414.Item(ItemVar)
                         Case "415"
                              valore = DB415.Item(ItemVar)
                         Case "416"
                              valore = DB416.Item(ItemVar)
                         Case "417"
                              valore = DB417.Item(ItemVar)
                         Case "418"
                              valore = DB418.Item(ItemVar)
                         Case "419"
                              valore = DB419.Item(ItemVar)
                         Case "420"
                              valore = DB420.Item(ItemVar)
                         Case "422"
                              valore = DB422.Item(ItemVar)
                         Case "424"
                              valore = DB424.Item(ItemVar)
                         Case "425"
                              valore = DB425.Item(ItemVar)
                         Case "426"
                              valore = DB426.Item(ItemVar)
                         Case "450"
                              valore = DB450.Item(ItemVar)
                         Case "460"
                              valore = DB460.Item(ItemVar)
                         Case "465"
                              valore = DB465.Item(ItemVar)
                         Case "470"
                              valore = DB470.Item(ItemVar)
                         Case "473"
                             valore = DB473.Item(ItemVar)
                         Case "480"
                             valore = DB480.Item(ItemVar)
                         Case "486"
                             valore = DB486.Item(ItemVar)
                        End Select
                        If IsNull(valore) = False And IsEmpty(valore) = False And valore <> "" Then
                           .Fields("Valore") = valore
                        Else
                           .Fields("Valore") = 0
                        End If
                        .Update
                End If
                 .MoveNext
               Wend
               .Close
               Set .ActiveConnection = Nothing
            End With
            Set rs = Nothing
            Set cn = Nothing
        Exit Sub
        
Errore:
             MsgBox "Controllare la mappa dati: elemento non caricato:" & IDitem & "," & ItemVar, vbCritical, "DP6.0"
             Exit Sub

errore2:
            MsgBox "Chiudere il database 'plc'", vbCritical, "DP6.0 - Files Copy"
             Exit Sub

End Sub

Sub CaricaDatiDBSimulazioneInPlc()
Dim IDitem, ItemVar As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
 
  cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\Plc.mdb;Persist Security Info=False"
  cn.Open
  With rs
       .Open "SELECT * FROM MappaDati", cn, adOpenKeyset, adLockReadOnly, adCmdText
       .MoveFirst
       While .EOF = False
          DoEvents
          If .Fields("Attivato") Then
                 IDitem = Left(.Fields("DBItem"), 3)
                 ItemVar = Right(.Fields("DBItem"), Len(.Fields("DBItem")) - 4)
                 Select Case IDitem
                 Case "400"
                      DB400.Item(ItemVar) = .Fields("Valore")
                 Case "402"
                      DB402.Item(ItemVar) = .Fields("Valore")
                 Case "403"
                      DB403.Item(ItemVar) = .Fields("Valore")
                 Case "410"
                      DB410.Item(ItemVar) = .Fields("Valore")
                 Case "411"
                      DB411.Item(ItemVar) = .Fields("Valore")
                 Case "412"
                      DB412.Item(ItemVar) = .Fields("Valore")
                 Case "413"
                      DB413.Item(ItemVar) = .Fields("Valore")
                 Case "414"
                      DB414.Item(ItemVar) = .Fields("Valore")
                 Case "415"
                      DB415.Item(ItemVar) = .Fields("Valore")
                 Case "416"
                      DB416.Item(ItemVar) = .Fields("Valore")
                 Case "417"
                      DB417.Item(ItemVar) = .Fields("Valore")
                 Case "418"
                      DB418.Item(ItemVar) = .Fields("Valore")
                 Case "419"
                      DB419.Item(ItemVar) = .Fields("Valore")
                 Case "420"
                      DB420.Item(ItemVar) = .Fields("Valore")
                 Case "422"
                      DB422.Item(ItemVar) = .Fields("Valore")
                 Case "424"
                      DB424.Item(ItemVar) = .Fields("Valore")
                 Case "425"
                      DB425.Item(ItemVar) = .Fields("Valore")
                 Case "426"
                      DB426.Item(ItemVar) = .Fields("Valore")
                 Case "450"
                      DB450.Item(ItemVar) = .Fields("Valore")
                 Case "460"
                      DB460.Item(ItemVar) = .Fields("Valore")
                 Case "465"
                      DB465.Item(ItemVar) = .Fields("Valore")
                 Case "470"
                      DB470.Item(ItemVar) = .Fields("Valore")
                 Case "473"
                      DB473.Item(ItemVar) = .Fields("Valore")
                 Case "480"
                      DB480.Item(ItemVar) = .Fields("Valore")
                 Case "486"
                      DB486.Item(ItemVar) = .Fields("Valore")
                End Select
         End If
         .MoveNext
       Wend
       .Close
       Set .ActiveConnection = Nothing
    End With
    Set rs = Nothing
    Set cn = Nothing
End Sub
Public Sub DownloadOrdiniEvasi()
   Dim i As Integer
   Dim MaxOrder As Integer
   Dim cn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Static UltimoCodice
   
   On Error Resume Next
   MaxOrder = Param.GetNumber("Par102_NumMaxOrdiniEvasi")
   If Param.GetNumber("Par102_NumMaxOrdiniEvasi") > 0 And UltimoCodice <> DB402.Word(18) Then
       cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
       cn.Open
       With rs
             .Open "SELECT * FROM OrdiniEvasi ORDER BY ID", cn, adOpenKeyset, adLockOptimistic, adCmdText
             .MoveLast
             UltimoCodice = DB402.Word(18)
             If UltimoCodice <> Val(.Fields("Ordine ID")) Then
                .AddNew
                .Fields("Ordine ID") = UltimoCodice
                .Fields("ID") = Date$ & "  " & Time$
                .Update
             End If
             If (.RecordCount) > MaxOrder - 1 Then
                   .MoveFirst
                   For i = 1 To (.RecordCount - MaxOrder)
                       .Delete adAffectCurrent
                       .MoveNext
                   Next
             End If
             .Close
             Set .ActiveConnection = Nothing
        End With
        Set rs = Nothing
        Set cn = Nothing
      End If
End Sub

Sub UpdateOrdineCorrente()
    Dim OrdineCorrente As New OrderClass
    
    OrdineCorrente.IDOrdine = DB470.Word(0)
    OrdineCorrente.UploadData DB470.Word(0)
    IDOrdineCorrente = OrdineCorrente.IDOrdine
    DescrOrdineCorrente = OrdineCorrente.Descrizione
    Set OrdineCorrente = Nothing
    
End Sub

'========================================================================
' carica in memoria le pagine
'========================================================================
    
Sub CaricaPagine(Optional ByVal SenzaPBar As Boolean)
    Const MAX_Pagine = 15
    
    If Not SenzaPBar Then
       SBstep = Int((frmPresentazione.PBar1.Max - frmPresentazione.PBar1.value) / MAX_Pagine)
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load Param
    Param.ScritteMultilingua
    Param.Barra21.Refresh_lingua
    
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load frmStampa

    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load frmStoccaggio
    frmStoccaggio.ScritteMultilingua
    frmStoccaggio.Barra21.Refresh_lingua

    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load BundlesLogForm
    BundlesLogForm.ScritteMultilingua
    BundlesLogForm.Barra21.Refresh_lingua
    
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load AlarmForm
    AlarmForm.ScritteMultilingua
    AlarmForm.Barra21.Refresh_lingua

    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load OrdersForm
    OrdersForm.ScritteMultilingua
    OrdersForm.Barra21.Refresh_lingua

    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load BundleForm
    BundleForm.ScritteMultilingua
    BundleForm.Barra21.Refresh_lingua

    If Not SenzaPBar Then
        frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
        frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
  '  Load frmSmussatrice
   ' frmSmussatrice.ScritteMultilingua
   '  frmSmussatrice.Barra21.Refresh_lingua
    
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
  '  Load FormWB
  '  FormWB.ScritteMultilingua
   ' FormWB.Barra21.Refresh_lingua
    
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
  '  Load FilettoForm
  '  FilettoForm.ScritteMultilingua
  '  FilettoForm.Barra21.Refresh_lingua
  
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load FormRegge
    FormRegge.ScritteMultilingua
    FormRegge.Barra21.Refresh_lingua
    
   If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
  '  Load FormLavaggio
  '  FormLavaggio.ScritteMultilingua

    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load CommandForm
    CommandForm.ScritteMultilingua
    CommandForm.Barra21.Refresh_lingua

    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load frmEntrata
    frmEntrata.ScritteMultilingua
    frmEntrata.Barra21.Refresh_lingua
    
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    Load frmStatistica
    frmStatistica.ScritteMultilingua
    frmStatistica.Barra21.Refresh_lingua
           
    If Not SenzaPBar Then
       frmPresentazione.PBar1.value = frmPresentazione.PBar1.value + SBstep
       frmPresentazione.PBar1.Text = Int(frmPresentazione.PBar1.value / PB_MAX * 100) & " %"
    End If
    
    If SenzaPBar Then
       frmAvvisi.ScritteMultilingua
       Unload frmAvvisi
       Set frmAvvisi = Nothing
       OrderDeleteForm.ScritteMultilingua
       RecipeModifyForm.ScritteMultilingua
       OrderModifyForm.ScritteMultilingua
    End If
End Sub

'================= SIMULAZIONE PLC VARIAZIONE DATI ===============================

Private Sub TSimula_Timer()
   Dim i As Integer
   Dim Nrandom As Integer
   Static Ciclo As Boolean
   Static Minuto As Variant
   Static NumtubiMin As Integer
   
   Ciclo = Not Ciclo
   DB402.Bit(0, 2) = True ' consenso cambio ordine
   If Boot = False Then
      DB403.Word(0) = 3 ' assegna i codici ordine alle varie zone
      DB450.Word(0) = 2
      DB465.Word(0) = 1
      DB460.Word(0) = 1
      DB470.Word(0) = 1
      DB473.Word(0) = 1
      DB480.Word(0) = 1
      DB486.Word(0) = 1
      DB410.Word(28) = 100 'tubificio
      DB450.Word(24) = 50
      DB450.Word(26) = 0
      DB450.Word(28) = 100
      Boot = True
   End If
   '  incrementa numero tubi
   DB420.Word(30) = DB420.Word(30) + Abs(Ciclo)
   '  incrementa numero pacco
   If DB420.Word(30) >= DB470.Word(22) Then DB420.Word(30) = 0: DB420.Word(32) = DB420.Word(32) + 1
   '   a 10000 pacchi si riazzera e riparte
   If DB420.Word(32) > 10000 Then DB420.Word(32) = 0
   ' animazione pacco in reggiatura
   If (DB422.Word(28) < DB480.Word(4) + 300) And DB480.Word(4) > 0 Then
      DB422.Word(28) = DB422.Word(28) + 12.5
   Else
      DB422.Word(28) = 0
   End If
   ' aggiorna il contatore tubi in entrata
   If Minuto <> Minute(Time) Then
      BufferTubiMin.MinuteActual = Minute(Time)
      BufferTubiMin.Data = Format(Date$, "mm-dd-yyyy") & "   " & Format(Time$, "hh.mm")
      For i = (BufferTubiMin.Minutelog - 1) To 1 Step -1
           BufferTubiMin.BufferData(i + 1, 1) = BufferTubiMin.BufferData(i, 1)
      Next
      Nrandom = (0.9 - (Int((9 * Rnd) + 1)) * 4 * (Abs((Int((9 * Rnd) + 1)) > 5) * (-1)) / 100)
      BufferTubiMin.BufferData(1, 1) = Int(NumtubiMin / 2 * Nrandom)
      NumtubiMin = 0
      BufferTubiMin.LetturaOK = True
      Minuto = Minute(Time)
   Else
      NumtubiMin = NumtubiMin + Abs(Ciclo)
   End If
   '  variazione pagina entrata
   DB410.Word(30) = (DB450.Word(24) * Abs(Int(Not (DB450.Bit(22, 0)))) + DB410.Word(28) * Abs(Int(DB450.Bit(22, 0)))) * (DB450.Word(28) / 100) + DB450.Word(26)
        
End Sub

Public Function RipristinaDatabase() As Boolean
   
    RipristinaDatabase = False
    ' connessione a database
    AdoLogPacco.CursorLocation = adUseClient
    AdoLogPacco.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\History.mdb;Persist Security Info=False"
    AdoLogPacco.CommandType = adCmdText
    AdoLogPacco.RecordSource = "select * from Bundles order by Contatore;"
    
    ' apertura database con eventuale ripristino del file danneggiato
    On Error GoTo Errore
    AdoLogPacco.Refresh
    AdoLogPacco.Recordset.ActiveConnection = Nothing
    Exit Function
Errore:
      '  If Err.Number <> 0 Then
            On Error GoTo 0
            On Error Resume Next
            
            RipristinaDatabase = True
            'If AdoLogPacco.Recordset.State = adStateOpen Then AdoLogPacco.Recordset.Close
            If FileEsistente("..\target\History.mdb") Then Kill "..\target\History.mdb"
            
            While FileEsistente("..\target\History.mdb")
               DoEvents
            Wend
            
            FileCopy "..\target\HistoryCopy.mdb", "..\target\History.mdb"
            
            While FileEsistente("..\target\History.mdb") = False
               DoEvents
            Wend
            
            Open "..\target\LogErrori.txt" For Append As #2
            Print #2, Format(Now, "dd-mm-yyyy hh:mm:ss") & " Restored History.mdb "
            Close #2
            Exit Function
      '  End If
End Function


Sub PrintLabel()

'    On Error Resume Next
'    Text1.Text = "Print(1," & Cartellino.Lingua & ",0)"
'    Text1.LinkPoke
'    If Err <> 0 Then
'        ================
'        Err.Clear
'        Text1.LinkTopic = "LabelProc|Form1"
'        Text1.LinkItem = "Text2"
'        Text1.LinkMode = vbLinkNone
'        Text1.LinkMode = vbLinkManual
'        Text1.LinkPoke
'        ================
'
'        If Err <> 0 Then
'            Shell App.path & "\LabelProc", vbNormalFocus
'            If Err <> 0 Then
'            ================
'            End If
'            Text1.LinkTopic = "LabelProc|Form1"
'            Text1.LinkItem = "Text2"
'            Text1.LinkMode = vbLinkNone
'            Text1.LinkMode = vbLinkManual
'            Text1.LinkPoke
'            ================
'        End If
'    End If
End Sub
Sub LetturaDatiPacco()
    Dim Ordine As OrderClass
    Dim i As Integer
    Dim MaxBundleLog As Long
    Dim FlagComapattazione As Boolean
    
    ' legge ordine da plc
    Set Ordine = New OrderClass
    Ordine.IDOrdine = DB402.Word(18)  'il codice dello stoccaggio si trova nel 402
    OrdersForm.LoadOrderData Ordine
    
    ' trasferimento sul cartellino dei dati dell'ordine
    Cartellino.Descrizione = Ordine.Descrizione
    For i = 1 To 10
        Cartellino.CampoManuale(i) = Ordine.CampoManuale(i)
    Next i
    ' aggiorna campo data e ora
    Cartellino.Data = Format(Now, "dd/mm/yyyy")
    Cartellino.Ora = Format(Now, "hh:mm")
    ' aggiorna altri campi automatici con dati da plc
    Cartellino.NumeroPacco = Format(DB402.Word(12), "###0")
    Cartellino.NumeroTubi = Format(DB402.Word(14), "###0")
    Cartellino.SpessoreTubo = Round(DB402.Word(28) / 10, 1) '
    Cartellino.DimensioniTubo = Round(DB402.Word(24) / 10) & "x" & Round(DB402.Word(26) / 10) & "x" & Format(Cartellino.SpessoreTubo, "###0.0")  ' altezza x larghezza  (decimi)
    Cartellino.LunghezzaTubo = Unit.m_To_Display_mm(DB402.Word(22) / 1000#)
        'peso del pacco
    If DB402.Word(16) >= 0 Then
       Cartellino.PesoPacco = DB402.Word(16)
    Else
      Cartellino.PesoPacco = "---"
    End If
    
    DB402.Bit(8, 0) = False
    
    Param.SetBit "Par204_AttivaPesa", True
    Set Ordine = Nothing
End Sub

Sub UploadDatiTurno()
   Dim cnn As New ADODB.Connection
   Dim rss As New ADODB.Recordset
   Dim StringaSql As String
   Dim i As Integer
   
          cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
          With rss
              StringaSql = "SELECT * FROM Turni"
              .Open StringaSql, cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
              For i = 1 To 3
                   DatiTurno(i, 1) = .Fields("TurnoInizio")
                   DatiTurno(i, 2) = .Fields("TurnoFine")
                   DatiTurno(i, 3) = .Fields("TurnoAlias")
                   .MoveNext
              Next
              .Close
              Set .ActiveConnection = Nothing
          End With
          Set rss = Nothing
          Set cnn = Nothing
End Sub
