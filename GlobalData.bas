Attribute VB_Name = "DataKERNELGlobal"
Option Explicit

'===============================================================
' oggetti per la connessione al database
'===============================================================
Public Enum UM
       mm = 0
       inch = 1
       kg = 2
       lB = 3
       ft = 4
       mt = 5
       mInch = 6
       yd = 7
       dmm = 8
       m_min = 9
       ft_min = 10
       m_s = 11
       inch_s = 12
       ft_s = 13
       inch_min = 14
       dm = 15
       dm3 = 16
       m3 = 17
       n = 18
       inch3 = 19
       kg_dm3 = 20
       n_dm3 = 21
       n_inch3 = 22
End Enum

Global PLC As New MairCOMS7.PLC_commands
Global tools As New MairCOMS7.Mair_Tools
Global lib As New MairCOMS7.Library

Public Conv_UM As New UMclass

Global S7_nome_collegamento As String
Global S7_nome_collegamentoI_O As String
Global OrderChanged As Boolean

Public Export As New ExportFileClass

Public Connessione As ADODB.Connection
Public RS_produzione As ADODB.Recordset
Public RS_AlarmsLOG As ADODB.Recordset
Public RS_LabelText As ADODB.Recordset
Public OrdiniMacchina As New OrderClass

' ==============================================================
'  dichiarazioni globali programma
' ==============================================================
Public DatiCartellino(20) As Integer
Public DatiTurno(3, 3) As Variant
Public AlarmLog As New LOGclass
Public RefreshListaOrdini As Boolean

Private Type TCal
        Day As String
        Month As String
        year As String
        Today As String
        Visible As Boolean
End Type

Public Calendario As TCal
Public TargetPath As String
Public BitmapPath As String
Public SourcePath As String
Public PrgPath As String
Public HelpPath As String
Global LogComPath As String
Global PrinterInstall As Boolean

Public WeightToPrint As String

Public PrintersList() As String
' costanti di uso comune

Public Enum Tubo
    Tondo = 1
    Quadro = 2
End Enum

Public Enum Pacco
    Esagono = 1
    Quadro = 2
End Enum

Public Enum Pennini
     TubiMin = 2
End Enum
' Struttura macchina

Public Const AbilitaPackPipe = 1
Public Const AbilitaRegge = 1
Public Const AbilitaSmusso = 0
Public Const AbilitaFiletto = 1
Public Const AbilitaStoccaggio = 1
Public Const AbilitaWB = 1
Public Const AbilitaCartellino = 0

'Colori stato macchina
Public Const ManualColor As Long = &H80FFFF   '&H90FF&  ' arancio
Public Const AutoColor As Long = &H80FF80   '&HFF00&    ' verde
Public Const SemiautoColor As Long = &HFF&  ' rosso
Public Const E_StopColor As Long = &HE0E0E0    ' grigio
' costanti per struttura dati pacco

Public Const MAX_ROWS   As Integer = 50  '  // numero massimo di file nel pacco e di tubi per fila
Public Const MAX_STRAPS As Integer = 12  '  // numero massimo di reggie nel pacco

'**************************************************
' DICHIARAZIONE DI OGGETTI GENERALI
' (creati e distrutti in frmKernel)
'**************************************************

' dati ordine futuro su pp
Public FutureOrder As OrderClass
' dati ordine attuale su pp
Public CurrentOrder As OrderClass
' dati ordine attuale in reggiatura
Public StrapOrder As OrderClass
' dati ultimo pacco pesato
'Public WeightOrder As OrderClass

' dati ordine in fase di modifica nella finestra OrderModifyForm
Public ModOrder As OrderClass
Public ModRecipe As RecipeClass
' dati ricetta in fase di modifica nella finestra RecipeModifyForm
Public Ricetta As RecipeClass
' dichiarazioni cartellino
Public Cartellino As LabelClass

' Dichiarazione oggetto unità di misura
Public Unit As UnitClass


'**************************************************
'**************************************************
' DEFINIZIONE DI COSTANTI GENERALI
'**************************************************
'**************************************************

' timeout conferma scrittura dati
Public Const MaxConfirmDelay As Integer = 1000

Public ModifyForm As OrderModifyForm

'********************************************************
' Definizione cartellino di stampa
'********************************************************
' 1) lunghezza dei campi variabili
Public Const PrintFieldLength As Integer = 30
' 2) stringa di default di un campo ad aggiornamento automatico
'    (deve essere lunga esattamente PrintFieldLength caratteri)
Public Const StringFieldLength As String = "012345678901234567890123456789"
' 3) Numero di campi
Public Const PrintFieldNumber As Integer = 33

' 6) carattere di separazione fra campi visualizzati
Public Const DisplaySeparator As String = " - "

'errore la dimensione del db è troppo piccola
Public ErrDBPiccolo As Boolean
Public NuovoAllarme As Boolean

'************************************************************************
'************************************************************************
'                             MAPPA DATI PLC
'************************************************************************
'************************************************************************
' DB di stato (manuale, automatico allarmi e conteggi)

'============ SOLA LETTURA: stato delle zone, allarmi ===================

Public DB400 As DBClass    ' stato linea (RO)
Public DB410 As DBClass    ' stato zona entrata (RO)
Public DB411 As DBClass    ' stato zona VR1 (RO)
Public DB412 As DBClass    ' stato zona bypass (RO)
Public DB413 As DBClass    ' stato zona entrata WB (RO)
Public DB414 As DBClass    ' stato zona uscita WB (RO)
Public DB415 As DBClass    ' stato zona WB (RO)
Public DB416 As DBClass    ' stato zona filettatura 1 (RO)
Public DB417 As DBClass    ' stato zona filettatura 2 (RO)
Public DB418 As DBClass    ' stato zona tappatrice (RO)
Public DB419 As DBClass    ' stato zona manicottatrice (RO)
Public DB420 As DBClass    ' stato zona MPS (RO)
Public DB422 As DBClass    ' stato zona trasportatori laterali (RO)
Public DB424 As DBClass    ' stato zona reggiatura e fasciatura (RO)
Public DB425 As DBClass    ' stato zona stoccaggio (RO)
Public DB426 As DBClass    ' stato zona stoccaggio finale(RO)
Public DB448 As DBClass    ' stato linea (RO)

'============ LETTURA/SCRITTURA: comunicazione,tracking ==================

'********* DB di tracking *************************************************

Public DB403 As DBClass     'commessa futura entrata
Public DB450 As DBClass     'zona entrata commessa attuale
Public DB451 As DBClass     'zona accumulo
Public DB452 As DBClass      'zona c
Public DB460 As DBClass      'zona WB
Public DB465 As DBClass     'zona CND
Public DB470 As DBClass     'dati locali 1mo polmone
Public DB471 As DBClass     'dati locali 2mo polmone
Public DB473 As DBClass     'dati locali mps
Public DB474 As DBClass     'dati locali Mensole
Public DB475 As DBClass     'dati locali carrello
Public DB480 As DBClass     'dati locali reggiatura
Public DB481 As DBClass     'dati locali catene
Public DB485 As DBClass     'dati locali catene
Public DB486 As DBClass     'dati locali stoccaggio

Public DBTestIN As DBClass
Public DBTestOUT As DBClass

'==============LETTURA/SCRITTURA: DB di comunicazione ===================

Public DB402 As DBClass    ' comandi vari per comunicazione


'********************** numero massimo di allarmi visualizzabili contemporaneamente
'numero massimo impostato di allarmi per area dati
Public Const MaxAllarmi As Integer = 100

'Allarmi pc (area interna)
Public ListaAllarmiPC(16) As String

'********************* Allarmi da plc
Public ListaAllarmiDB400(MaxAllarmi) As String
Public ListaAllarmiDB410(MaxAllarmi) As String
Public ListaAllarmiDB411(MaxAllarmi) As String
Public ListaAllarmiDB412(MaxAllarmi) As String
Public ListaAllarmiDB413(MaxAllarmi) As String
Public ListaAllarmiDB414(MaxAllarmi) As String
Public ListaAllarmiDB415(MaxAllarmi) As String
Public ListaAllarmiDB416(MaxAllarmi) As String
Public ListaAllarmiDB417(MaxAllarmi) As String
Public ListaAllarmiDB418(MaxAllarmi) As String
Public ListaAllarmiDB419(MaxAllarmi) As String
Public ListaAllarmiDB420(MaxAllarmi) As String
Public ListaAllarmiDB422(MaxAllarmi) As String
Public ListaAllarmiDB424(MaxAllarmi) As String
Public ListaAllarmiDB425(MaxAllarmi) As String
Public ListaAllarmiDB426(MaxAllarmi) As String

'********************* Messaggi da plc

Public ListaMessaggiDB400(MaxAllarmi) As String
Public ListaMessaggiDB410(MaxAllarmi) As String
Public ListaMessaggiDB411(MaxAllarmi) As String
Public ListaMessaggiDB412(MaxAllarmi) As String
Public ListaMessaggiDB413(MaxAllarmi) As String
Public ListaMessaggiDB414(MaxAllarmi) As String
Public ListaMessaggiDB415(MaxAllarmi) As String
Public ListaMessaggiDB416(MaxAllarmi) As String
Public ListaMessaggiDB417(MaxAllarmi) As String
Public ListaMessaggiDB418(MaxAllarmi) As String
Public ListaMessaggiDB419(MaxAllarmi) As String
Public ListaMessaggiDB420(MaxAllarmi) As String
Public ListaMessaggiDB422(MaxAllarmi) As String
Public ListaMessaggiDB424(MaxAllarmi) As String
Public ListaMessaggiDB425(MaxAllarmi) As String
Public ListaMessaggiDB426(MaxAllarmi) As String

'************************ Eco allarmi da plc

Public EcoAllarmiDB400(26, 8) As Boolean
Public EcoAllarmiDB410(26, 8) As Boolean
Public EcoAllarmiDB411(26, 8) As Boolean
Public EcoAllarmiDB412(26, 8) As Boolean
Public EcoAllarmiDB413(26, 8) As Boolean
Public EcoAllarmiDB414(26, 8) As Boolean
Public EcoAllarmiDB415(26, 8) As Boolean
Public EcoAllarmiDB416(26, 8) As Boolean
Public EcoAllarmiDB417(26, 8) As Boolean
Public EcoAllarmiDB418(26, 8) As Boolean
Public EcoAllarmiDB419(26, 8) As Boolean
Public EcoAllarmiDB420(26, 8) As Boolean
Public EcoAllarmiDB422(26, 8) As Boolean
Public EcoAllarmiDB424(26, 8) As Boolean
Public EcoAllarmiDB425(26, 8) As Boolean
Public EcoAllarmiDB426(26, 8) As Boolean

' ==================== Dati intestazione pagine locali =========================
Public Type TPaginaPacco
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
Public Type TPaginaReggiatura
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
Public Type TPaginaStoccaggio
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
Public Type TPaginaEntrata
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
Public Type TPaginaWB
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
Public Type TPaginaLav
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
Public Type TPaginaArchivia
                  Ordine_Descrizione As String
                  Ricetta_Descrizione As String
End Type
' buffer dati grafico tubi/min
Public Type TTubiMin
                  Data As String
                  LetturaOK As Boolean
                  MinuteActual As Integer
                  Minutelog As Integer
                  BufferData() As Integer
End Type

Public NomeProfili(3) As String
Public PaginaPacco As TPaginaPacco
Public PaginaReggiatura As TPaginaReggiatura
Public PaginaStoccaggio As TPaginaStoccaggio
Public PaginaEntrata As TPaginaEntrata
Public PaginaWb As TPaginaWB
Public PaginaLav As TPaginaLav
Public PaginaArchivia As TPaginaArchivia
Public BufferTubiMin As TTubiMin

'********************* commessa futura in entrata DB403

Public Sub DownloadDB403(OrderSource As OrderClass, RecipeSource As RecipeClass)
    On Error Resume Next    ' per eventuali overflow
        DB403.Word(0) = OrderSource.IDOrdine
        If RecipeSource.TipoTubo = Tubo.Tondo Then
            DB403.Bit(2, 0) = True
        Else
            DB403.Bit(2, 0) = False
        End If
        If Param.GetBit("Par201_AbilitazioneProfili") Then
           DB403.Bit(2, 1) = RecipeSource.Profilo And 1
           DB403.Bit(2, 2) = RecipeSource.Profilo And 2
           DB403.Bit(2, 3) = RecipeSource.Profilo And 3
        End If
        DB403.Word(4) = RecipeSource.TuboLunghezza * 1000#
        DB403.Word(6) = RecipeSource.TuboLarghezza * 10000#
        DB403.Word(8) = RecipeSource.TuboAltezza * 10000#
        DB403.Word(10) = RecipeSource.TuboSpessore * 10000#
        DB403.Word(12) = RecipeSource.TuboPesoTeorico * 100#
    On Error GoTo 0
    DB403.Bit(20, 0) = True      ' dati modificati
End Sub

'' zona accumulo
'Public Sub DownloadDB451(OrderSource As OrderClass)
'    DB451.Word(0) = OrderSource.IDOrdine
'    ScriviStringa DB451, 30, OrderSource.Descrizione
'End Sub

Public Sub DownloadDB450(RecipeSource As RecipeClass)
        If RecipeSource.TipoTubo = Tubo.Tondo Then
            DB450.Bit(2, 0) = True
        Else
            DB450.Bit(2, 0) = False
        End If
        If Param.GetBit("Par201_AbilitazioneProfili") Then
           DB450.Bit(2, 1) = RecipeSource.Profilo And 1
           DB450.Bit(2, 2) = RecipeSource.Profilo And 2
           DB450.Bit(2, 3) = RecipeSource.Profilo And 3
        End If
        DB450.Word(4) = RecipeSource.TuboLunghezza * 1000#
        DB450.Word(6) = RecipeSource.TuboLarghezza * 10000#
        DB450.Word(8) = RecipeSource.TuboAltezza * 10000#
        DB450.Word(10) = RecipeSource.TuboSpessore * 10000#
        DB450.Word(12) = RecipeSource.TuboPesoTeorico * 100#
        
'   DB450.Bit(22, 0) = RecipeSource.Bypass0
'   DB450.Bit(22, 1) = RecipeSource.Bypass1
'   DB450.Bit(22, 2) = RecipeSource.Bypass2
'   DB450.Bit(22, 3) = RecipeSource.Bypass3
'   DB450.Word(24) = RecipeSource.VelVR1
'   DB450.Word(26) = RecipeSource.VelVR2
'   DB450.Word(28) = RecipeSource.VelMB1
'   DB450.Word(30) = RecipeSource.VelMB2
End Sub

''zona MC
'Public Sub DownloadDB450(OrderSource As OrderClass)
'    DB450.Word(0) = OrderSource.IDOrdine
'    On Error Resume Next    ' per eventuali overflow
'        If OrderSource.LunghezzaTagliata < OrderSource.LunghezzaBarra Then
'            DB450.DWord(426) = OrderSource.LunghezzaTagliata * 1000#
'        Else
'            DB450.DWord(426) = 0
'        End If
'    On Error GoTo 0
'    ScriviStringa DB450, 460, OrderSource.Descrizione
'    ScriviStringa DB450, 480, OrderSource.TipoMateriale
'End Sub

Public Sub DownloadDB460(RecipeSource As RecipeClass)
        If RecipeSource.TipoTubo = Tubo.Tondo Then
            DB460.Bit(2, 0) = True
        Else
            DB460.Bit(2, 0) = False
        End If
        If Param.GetBit("Par201_AbilitazioneProfili") Then
           DB460.Bit(2, 1) = RecipeSource.Profilo And 1
           DB460.Bit(2, 2) = RecipeSource.Profilo And 2
           DB460.Bit(2, 3) = RecipeSource.Profilo And 3
        End If
        DB460.Word(4) = RecipeSource.TuboLunghezza * 1000#
        DB460.Word(6) = RecipeSource.TuboLarghezza * 10000#
        DB460.Word(8) = RecipeSource.TuboAltezza * 10000#
        DB460.Word(10) = RecipeSource.TuboSpessore * 10000#
        DB460.Word(12) = RecipeSource.TuboPesoTeorico * 100#
        DB460.Word(34) = RecipeSource.VelMB3
End Sub

'' zona smusso
'Public Sub DownloadDB460(OrderSource As OrderClass)
'    DB460.Word(0) = OrderSource.IDOrdine
'
'    On Error Resume Next    ' per eventuali overflow
'        DB460.Word(4) = OrderSource.LunghezzaTagliata * 1000#
'        DB460.Word(12) = OrderSource.PesoBarraAlMetro * 100# * OrderSource.LunghezzaTagliata
'
'        DB460.DWord(52) = OrderSource.LunghezzaSmusso * 1000000#  ' smusso 1
'        DB460.DWord(92) = OrderSource.LunghezzaSmusso * 1000000#  ' smusso 2
'
'        ' inserimento automatico smussatura
'        If DB460.DWord(52) > 0 Then
'            DB460.Bit(78, 1) = True
'            DB460.Bit(118, 1) = True
'        Else
'            DB460.Bit(78, 1) = False
'            DB460.Bit(118, 1) = False
'        End If
'
'    On Error GoTo 0
'    ScriviStringa DB460, 120, OrderSource.Descrizione
'    ScriviStringa DB460, 140, OrderSource.TipoMateriale
'    ScriviStringa DB460, 160, OrderSource.AngoloSmusso
'End Sub



'Public Sub DownloadDB486(OrderSource As OrderClass)
'    DB486.Word(0) = OrderSource.IDOrdine
''    ScriviStringa DB486, 84, OrderSource.Descrizione
'End Sub


' scrittura dei dati su DB470
Public Sub DownloadDB470(OrderSource As OrderClass, RecipeSource As RecipeClass)
    Dim i As Integer
    On Error Resume Next    ' per eventuali overflow
        DB470.Word(0) = OrderSource.IDOrdine ' scarica il codice ordine nella DB470
        If RecipeSource.TipoTubo = Tubo.Tondo Then
            DB470.Bit(2, 0) = True
        Else
            DB470.Bit(2, 0) = False
        End If
        If Param.GetBit("Par201_AbilitazioneProfili") Then
           DB470.Bit(2, 1) = RecipeSource.Profilo And 1
           DB470.Bit(2, 2) = RecipeSource.Profilo And 2
           DB470.Bit(2, 3) = RecipeSource.Profilo And 3
        End If
        DB470.Word(4) = RecipeSource.TuboLunghezza * 1000#
        DB470.Word(6) = RecipeSource.TuboLarghezza * 10000#
        DB470.Word(8) = RecipeSource.TuboAltezza * 10000#
        DB470.Word(10) = RecipeSource.TuboSpessore * 10000#
        DB470.Word(12) = RecipeSource.TuboPesoTeorico * 100#
        If RecipeSource.TipoPacco = Pacco.Esagono Then
           DB470.Bit(20, 0) = True
        Else
           DB470.Bit(20, 0) = False
        End If
        DB470.Word(22) = RecipeSource.NumeroTubiPacco
        DB470.Word(24) = RecipeSource.PaccoPesoTeorico
        
        If RecipeSource.TipoPacco = Pacco.Esagono Then
            DB470.Word(26) = RecipeSource.PaccoLarghezzaBaseEsagono * 1000#
        Else
            DB470.Word(26) = RecipeSource.TuboLarghezza * RecipeSource.TubiFila(1) * 1000#
        End If
        DB470.Word(28) = RecipeSource.PaccoLarghezza * 1000#
        DB470.Word(30) = RecipeSource.PaccoLarghezzaLatoEsagono * 1000#
        DB470.Word(32) = RecipeSource.PaccoAltezza * 1000#
        DB470.Word(34) = RecipeSource.FilaUscitaControsagoma
        For i = 1 To 50
            DB470.Word(78 + i * 2) = RecipeSource.TubiFila(i)
        Next i
        If OrderSource.ModoCambioOrdine = 1 Then
            ' cambio ordine a fine pacchi
            DB402.Bit(0, 5) = False
            ' arresto a fine pacchi
            DB402.Bit(0, 6) = True
        Else
            If OrderSource.ModoCambioOrdine = 2 Then
                ' cambio ordine a fine pacchi
                DB402.Bit(0, 5) = True
                ' arresto a fine pacchi
                DB402.Bit(0, 6) = False
            Else
                ' cambio ordine a fine pacchi
                DB402.Bit(0, 5) = False
                ' arresto a fine pacchi
                DB402.Bit(0, 6) = False
            End If
        End If
        DB402.Word(2) = OrderSource.PresetPacchi
        DB470.Word(66) = RecipeSource.VelTR     'scarica la velocita del tr sal
'        If Param.GetBit("Par229_STORAGE_DEST") Then
        DB470.Word(36) = RecipeSource.Destination
'        End If
        DB470.Bit(40, 0) = True ' dati modificati
    On Error GoTo 0
End Sub

Public Sub DownloadDB473(RecipeDest As RecipeClass)
   
   DB473.Word(64) = RecipeDest.VelMPS
   
End Sub


' lettura dei dati da DB480
Public Sub UploadDB480(RecipeDest As RecipeClass)
    Dim i As Integer
    On Error Resume Next    ' per eventuali overflow
        For i = 1 To 12
            RecipeDest.QuotaReggia(i) = DB480.Word(64 + i * 2) / 1000#
        Next i
        RecipeDest.NumeroRegge = DB480.Word(64)
        RecipeDest.Regg1 = DB422.Bit(27, 2)
        RecipeDest.Regg2 = DB422.Bit(27, 3)
    On Error GoTo 0
End Sub
' scrittura dei dati su DB480
Public Sub DownloadDB480(RecipeSource As RecipeClass)
    Dim i As Integer
    On Error Resume Next    ' per eventuali overflow
        For i = 1 To 12
            DB480.Word(64 + i * 2) = RecipeSource.QuotaReggia(i) * 1000#
        Next i
        DB480.Word(4) = RecipeSource.TuboLunghezza * 1000
        DB480.Word(64) = RecipeSource.NumeroRegge
        If RecipeSource.Regg1 = False And RecipeSource.Regg2 = False Then
           DB422.Bit(27, 2) = True
        Else
           DB422.Bit(27, 2) = RecipeSource.Regg1
        End If
        DB422.Bit(27, 3) = RecipeSource.Regg2
       
    On Error GoTo 0
End Sub

Public Sub DownloadD_ricetta_B480(RecipeSource As RecipeClass)
    Dim i As Integer
    
    On Error Resume Next    ' per eventuali overflow
        For i = 1 To 12
            DB480.Word(64 + i * 2) = RecipeSource.QuotaReggia(i) * 1000#
        Next i
        DB480.Word(4) = RecipeSource.TuboLunghezza * 1000
        DB480.Word(64) = RecipeSource.NumeroRegge
     
    On Error GoTo 0
End Sub

Function FileEsistente(NomeFile As String) As Boolean
    On Error GoTo GestoreErrori
    'Controlla l'esistenza del file specificato.
    If NomeFile <> "" Then
        FileEsistente = IIf(Dir(NomeFile, vbNormal Or vbHidden Or vbReadOnly Or vbSystem) <> "", True, False)
    Else
        FileEsistente = False
    End If
    Exit Function
GestoreErrori:
    'Sì è verificato un errore: il file non esiste.
    FileEsistente = False
End Function
Function CartellaEsistente(NomeDir As String) As Boolean
    On Error GoTo GestoreErrori
    'Controlla l'esistenza della cartella specificata.
    If NomeDir <> "" Then
        CartellaEsistente = IIf(Dir(NomeDir, vbDirectory Or vbNormal Or vbReadOnly Or vbSystem Or vbHidden) <> "", True, False)
    Else
        CartellaEsistente = False
    End If
    Exit Function
GestoreErrori:
    'Sì è verificato un errore: il file non esiste.
    CartellaEsistente = False
End Function

Sub StartupDatiPagina()
   Dim Ordine As New OrderClass

   Ordine.UploadData frmKernel.CodOrdineCorrente.CodPacco 'upload pacco
   PaginaPacco.Ordine_Descrizione = Ordine.Descrizione
   PaginaPacco.Ricetta_Descrizione = Ordine.IDRicetta
   Ordine.UploadData frmKernel.CodOrdineCorrente.CodRegge 'upload regge
   PaginaReggiatura.Ordine_Descrizione = Ordine.Descrizione
   PaginaReggiatura.Ricetta_Descrizione = Ordine.IDRicetta
   Ordine.UploadData frmKernel.CodOrdineCorrente.CodStoccaggio 'upload stoccaggio
   PaginaStoccaggio.Ordine_Descrizione = Ordine.Descrizione
   PaginaStoccaggio.Ricetta_Descrizione = Ordine.IDRicetta
   Ordine.UploadData frmKernel.CodOrdineCorrente.Cod402 ' ' upload pc-plc
   PaginaArchivia.Ordine_Descrizione = Ordine.Descrizione
   PaginaArchivia.Ricetta_Descrizione = Ordine.IDRicetta
   Ordine.UploadData frmKernel.CodOrdineCorrente.CodEntrata 'upload entrata
   PaginaEntrata.Ordine_Descrizione = Ordine.Descrizione
   PaginaEntrata.Ricetta_Descrizione = Ordine.IDRicetta
   Ordine.UploadData frmKernel.CodOrdineCorrente.CodWB  'upload WB-filettatrici
   PaginaWb.Ordine_Descrizione = Ordine.Descrizione
   PaginaWb.Ricetta_Descrizione = Ordine.IDRicetta
   Set Ordine = Nothing
End Sub

' *********************** definizione oggetto grafico **************************
  
Sub InizializzaOggettoGrafico(ByVal Oggetto As MSChart)
   Dim i As Integer
   Dim LabelTempo As String
           
   With Oggetto
           
           '------------------------------------------------------------------------------------------------------------------
           ' caratteristiche grafico
           ' Caratteristiche titolo
           '------------------------------------------------------------------------------------------------------------------
           .Title.Text = Param.Text("Vlinea")
           .Title.VtFont.Style = VtFontStyle.VtFontStyleBold
           .Title.VtFont.Size = 12
           .Title.VtFont.Name = "Arial"
           .Title.VtFont.VtColor.Set 0, 255, 0
           'caratteristiche grafico
            .chartType = VtChChartType2dLine
            .Backdrop.Frame.FrameColor.Set 0, 0, 0
            .Backdrop.Frame.SpaceColor.Set 255, 0, 0
            .Backdrop.Shadow.Style = VtShadowStyleNull
            'caratteristiche leggenda
            .ShowLegend = False
            With .Legend
                    .VtFont.Name = "Arial"
                    .VtFont.Size = 12
            End With
            'caratteristiche disegno
            With .Plot
                   .AutoLayout = True
                   ' collezione pennini
                   With .SeriesCollection(1)  'selezionato il pennino 1 dell'insieme
                           .LegendText = "Tubi/min"
                           .SeriesType = VtChSeriesType2dLine
                           .pen.VtColor.Set 0, 255, 0
                   End With
                   '------------------------------------------------------------------------------------------------------------------
                   ' Caratteristiche ASSE X
                   '------------------------------------------------------------------------------------------------------------------
                   With .Axis(VtChAxisIdX)     ' Riferimento all'asse x
                           .AxisTitle = Param.Text("Tempo")                                 'titolo asse
                           '.AxisTitle.VtFont.Effect = VtFontEffect.VtFontEffectUnderline
                           .AxisTitle.VtFont.Style = VtFontStyle.VtFontStyleBold
                           .AxisTitle.VtFont.Name = "Arial"
                           .AxisTitle.VtFont.Size = 12
                           .AxisGrid.MajorPen.VtColor.Set 0, 0, 120   'colore divisione
                           .AxisScale.Type = VtChScaleTypeLinear    'tipo scala lineare
                           .ValueScale.Auto = False                          'disattiva la scalatura automatica dei valori
                           .CategoryScale.DivisionsPerTick = 1           'passo tra le tacche
                           .CategoryScale.DivisionsPerLabel = 1         'passo tra le etichette
                           .Labels(1).Auto = False
                           .Labels(1).TextLayout.Orientation = VtOrientationHorizontal
                           .CategoryScale.LabelTick = True
                           .Tick.Length = 100
                           .Tick.Style = VtChAxisTickStyleOutside
                   End With
                   
                   '------------------------------------------------------------------------------------------------------------------
                   ' Caratteristiche ASSE Y
                   '------------------------------------------------------------------------------------------------------------------
                
                   With .Axis(VtChAxisIdY)      ' Riferimento all'asse y
                           .AxisTitle.Text = Param.Text("Tubi")
                           .AxisTitle.VtFont.Style = VtFontStyle.VtFontStyleBold
                           .AxisTitle.VtFont.Name = "Arial"
                           .AxisTitle.VtFont.Size = 12
                           .AxisGrid.MajorPen.Style = VtPenStyleDitted
                           .AxisGrid.MajorPen.VtColor.Set 0, 255, 255
                           .ValueScale.Auto = True
                           .ValueScale.Maximum = 100
                           .ValueScale.Minimum = 0
                           .ValueScale.MajorDivision = 10
                           .ValueScale.Minimum = 0
                   End With
           End With
           
           ' definizione massimi
           BufferTubiMin.Minutelog = 15
           
           ReDim BufferTubiMin.BufferData(1 To BufferTubiMin.Minutelog, 1 To Pennini.TubiMin)
           BufferTubiMin.MinuteActual = Minute(Time)
           
           ' carica valori di inizializzazione
           .ChartData = BufferTubiMin.BufferData
           For i = 1 To BufferTubiMin.Minutelog
              .Row = i
              LabelTempo = Format(TimeSerial(Hour(Time), Minute(Time) - i + 1, Second(Time)), "hh.mm")
              .RowLabel = LabelTempo
           Next
           .RowCount = 15
           .ColumnCount = 1
           BufferTubiMin.LetturaOK = False
   End With
End Sub

Sub DownloadDatiGrafico(ByVal Oggetto As MSChart)
     Dim i As Integer
     Dim LabelTempo As String
     
     With Oggetto
           .ChartData = BufferTubiMin.BufferData
           For i = 1 To BufferTubiMin.Minutelog
                .Row = i
                LabelTempo = Format(TimeSerial(Hour(Time), Minute(Time) - i + 1, Second(Time)), "hh.mm")
                .RowLabel = LabelTempo
           Next
           .RowCount = 15
           .ColumnCount = 1
           BufferTubiMin.LetturaOK = False
      End With
End Sub

Sub Main()
   Dim x As Printer
   Dim RS_Param As ADODB.Recordset
   
   ChDir (App.path)
   PrgPath = CurDir & "\"
   ChDir ("..")
   TargetPath = CurDir & "\Target\"
   SourcePath = CurDir & "\Source\"
   BitmapPath = CurDir & "\Bitmap\"
   HelpPath = CurDir & "\Help\"
   LogComPath = CurDir & "\Target\PagineHelp\Errori.htm"
   ChDir (PrgPath)
  
   '===============================================================================
    '===============================================================================
    ' apertura database produzione
    '===============================================================================
    '===============================================================================
    
    On Error GoTo Errore
    Set RS_Param = New ADODB.Recordset
    If OrdiniMacchina.Client_OPEN_DBproduzione Then
       If OrdiniMacchina.Client_OrdiniInMacchina_CreateRS Then
          OrdiniMacchina.Client_OrdiniInMacchina_Refresh True
          DoEvents
          On Error Resume Next
          OrdiniMacchina.Client_CLOSE_RSproduzione
          OrdiniMacchina.Client_CLOSE_DBproduzione
          Kill "..\target\productionCopy.mdb"
          FileCopy "..\target\production.mdb", "..\target\productionCopy.mdb"
          DoEvents
          CompattazioneDB "..\target\production.mdb"
          DoEvents
          On Error GoTo Errore
          If OrdiniMacchina.Client_OPEN_DBproduzione Then
             If OrdiniMacchina.Client_OrdiniInMacchina_CreateRS Then
                OrdiniMacchina.Client_OrdiniInMacchina_Refresh True
             End If
          End If
          On Error Resume Next
          RS_Param.Open "Select * from Connections WHERE Enable=True", Connessione, adOpenKeyset, adLockReadOnly
          If RS_Param.EOF Then
             MsgBox "Select the connection type (DB production.mdb/connections)", vbCritical, "Dp6": End
          End If
          RS_Param.MoveFirst
          S7_nome_collegamento = RS_Param("Connection") & "DB"
          S7_nome_collegamentoI_O = RS_Param("Connection")
          RS_Param.Close
          Set RS_Param = Nothing
       End If
    Else
       Err.Raise 70
    End If
    On Error GoTo 0
    '===================================================
'   On Error Resume Next
'   Shell App.path & "\LabelProc", vbNormalFocus
'   If Err <> 0 Then
'      MsgBox "LabelProc not present:install it for print the ticket", vbInformation, "DP6 - Starter"
'   End If
'   On Error GoTo 0
   '===================================================
   
    On Error Resume Next
    '===================================================
    If AlarmLog.TestFile(TargetPath & "Alarmslog.txt", 11, 10, 5, 1, 1) = False Then
       AlarmLog.GestioneErrore (TargetPath & "Alarmslog.txt")
    End If
    '===================================================
   
   PrinterInstall = False
   For Each x In Printers
      If InStr(x.DeviceName, "EasyCoder") <> 0 Then
         'If X.Orientation = vbPRORPortrait Then
         ' Imposta la stampante come predefinita di sistema.
         Set Printer = x
         PrinterInstall = True
         ' Interrompe la ricerca di una stampante.
         Exit For
      End If
   Next
   PrintersListRefresh
   frmSplash.Show
   Exit Sub
   
Errore:
    OrdiniMacchina.Client_CLOSE_DBproduzione
    Kill "..\target\production.mdb"
    FileCopy "..\target\productionCopy.mdb", "..\target\production.mdb"
    End
End Sub

Sub PrintersListRefresh()
   Dim x As Printer
   Dim i As Integer
   
   ReDim PrintersList(0)
   PrinterInstall = False
   i = 0
   For Each x In Printers
      If InStr(x.DeviceName, "EasyCoder") <> 0 Then
         PrinterInstall = True
         DoEvents
      End If
      i = i + 1
      ReDim Preserve PrintersList(i)
      PrintersList(i) = x.DeviceName
   Next
End Sub
Sub CompattazioneDB(DataBaseName As String)
    Dim TmpName As String
'    Dim MyJro As jro.JetEngine ' per access2000 nei riferimenti includere msjro.dll (microsoft jet and replication object 2.5 o 2.6)
    

    TmpName = "Tmp.mdb"
    On Error Resume Next
        '1) libera spazio per un file temporaneo
        If Dir(TmpName) <> "" Then Kill TmpName
        '2) Compatta il database in un file temporaneo
        ' access 97
        DBEngine.CompactDatabase DataBaseName, TmpName
        ' access 2000
'        Set MyJro = New jro.JetEngine                            'DB Sorgente                                          'DB Destinazione
'        MyJro.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DataBaseName & ";", _
'                              "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & TmpName & ";" & _
'                              "Jet OLEDB:Engine Type=5"
        
        If Err.number <> 0 Then
            'MsgBox "ERRORE DI COMPATTAZIONE !", vbInformation, "Warning"
           ' PcAlarm(ErrCompattDataBase) = True
        Else
            '3) Ripristina il database con il nome originale
            Kill DataBaseName
            Name TmpName As DataBaseName
            If Err.number <> 0 Then
                'MsgBox "ERRORE DI RIPRISTINO DATABASE !", vbInformation, "Warning"
                'PcAlarm(ErrRipristDataBase) = True
            Else
                'MsgBox "Tutto Bene", vbOKOnly
                'PcAlarm(ErrCompattDataBase) = False
                'PcAlarm(ErrRipristDataBase) = False
            End If
        End If
    On Error GoTo 0
End Sub

