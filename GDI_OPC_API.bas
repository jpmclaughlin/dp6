Attribute VB_Name = "GDI_OPC_API"
Option Explicit


'dichiarazione API: acquisizione tasto hook di tastiera

' tipi per controllo mouse e rettangolo windows

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

' costanti API per il GDI tastiera e mouse

Public Const BUTTON_NONE = 0
Public Const BUTTON_UP = 1
Public Const BUTTON_DOWN = 2

Public Const BACKGROUND_COLOR = &H80000010
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8

Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8

Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const BF_DIAGONAL = &H10


Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP _
             Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM _
             Or BF_RIGHT)

Public Const BF_MIDDLE = &H800
Public Const BF_SOFT = &H1000
Public Const BF_ADJUST = &H2000
Public Const BF_FLAT = &H4000
Public Const BF_MONO = &H8000
Global Const English = &H409
Global Const OPC_DS_CACHE = 1
Global Const OPC_DS_DEVICE = 2
Global Const ADVICE_START = 100
Global Const WRITEASYNC_ID = 1
Global Const READASYNC_ID = 2
Global Const REFRESHASYNC_ID = 3
'----------------------------------------------------------------------------
'Constante  Qualities dei dati OPC
'----------------------------------------------------------------------------
Global Const OPC_QUALITY_MASK = &HC0
Global Const OPC_STATUS_MASK = &HFC
Global Const OPC_LIMIT_MASK = &H3
Global Const OPC_QUALITY_BAD = &H0
Global Const OPC_QUALITY_UNCERTAIN = &H40
Global Const OPC_QUALITY_GOOD = &HC0
Global Const OPC_QUALITY_CONFIG_ERROR = &H4
Global Const OPC_QUALITY_NOT_CONNECTED = &H8
Global Const OPC_QUALITY_DEVICE_FAILURE = &HC
Global Const OPC_QUALITY_SENSOR_FAILURE = &H10
Global Const OPC_QUALITY_LAST_KNOWN = &H14
Global Const OPC_QUALITY_COMM_FAILURE = &H18
Global Const OPC_QUALITY_OUT_OF_SERVICE = &H1C
Global Const OPC_QUALITY_LAST_USABLE = &H84
Global Const OPC_QUALITY_SENSOR_CAL = &H90
Global Const OPC_QUALITY_EGU_EXCEEDED = &H94
Global Const OPC_QUALITY_SUB_NORMAL = &H98
Global Const OPC_QUALITY_LOCAL_OVERRIDE = &HD8

'======================== dichiarazioni funzioni api ===============================

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function PtInRect Lib "user32" (RECT As RECT, ByVal lPtX As Long, ByVal lPtY As Long) As Integer
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' variabili d'ambiente

Global lpPoint As POINTAPI
Global OggServer As OPCServer

'====================================================================================
'                             INFORMAZIONI LOCALI
'====================================================================================
'================================================================================
' GESTIONE LOCALI
'================================================================================
 Const LOCALE_ITIME = &H23        '  indicatore formato di ora
 Const LOCALE_STIMEFORMAT = &H1003      '  stringa formato ora
 Const LOCALE_STIME = &H1E        '  separatore di ora
 Const LOCALE_STHOUSAND = &HF         '  separatore delle migliaia
 Const LOCALE_SSHORTDATE = &H1F        '  stringa formato data breve
 Const LOCALE_SPOSITIVESIGN = &H50        '  segno positivo
 Const LOCALE_SNEGATIVESIGN = &H51        '  segno negativo
 Const LOCALE_SNATIVELANGNAME = &H4         '  nome nativo della lingua *********
 Const LOCALE_SNATIVEDIGITS = &H13        '  ascii nativo da 0 a 9
 Const LOCALE_SNATIVECTRYNAME = &H8         '  nome nativo del paese
 Const LOCALE_SMONTHOUSANDSEP = &H17        '  separatore delle migliaia della valuta
 Const LOCALE_SMONTHNAME9 = &H40        '  nome lungo per settembre
 Const LOCALE_SMONTHNAME8 = &H3F        '  nome lungo per agosto
 Const LOCALE_SMONTHNAME7 = &H3E        '  nome lungo per luglio
 Const LOCALE_SMONTHNAME6 = &H3D        '  nome lungo per giugno
 Const LOCALE_SMONTHNAME5 = &H3C        '  nome lungo per maggio
 Const LOCALE_SMONTHNAME4 = &H3B        '  nome lungo per aprile
 Const LOCALE_SMONTHNAME3 = &H3A        '  nome lungo per marzo
 Const LOCALE_SMONTHNAME2 = &H39        '  nome lungo per febbraio
 Const LOCALE_SMONTHNAME12 = &H43        '  nome lungo per dicembre
 Const LOCALE_SMONTHNAME11 = &H42        '  nome lungo per novembre
 Const LOCALE_SMONTHNAME10 = &H41        '  nome lungo per ottobre
 Const LOCALE_SMONTHNAME1 = &H38        '  nome lungo per gennaio
 Const LOCALE_SMONGROUPING = &H18        '  raggruppamento valuta
 Const LOCALE_SMONDECIMALSEP = &H16        '  separatore decimale valuta
 Const LOCALE_SLONGDATE = &H20              '  stringa formato data lungo
 Const LOCALE_SLIST = &HC                          '  separatore elenchi **************
 Const LOCALE_SLANGUAGE = &H2               '  nome localizzato della lingua
 Const LOCALE_SINTLSYMBOL = &H15           '  simbolo internazionale valuta
 Const LOCALE_SGROUPING = &H10              '  raggruppamento cifre
 Const LOCALE_SENGLANGUAGE = &H1001    '  nome in inglese della lingua
 Const LOCALE_SENGCOUNTRY = &H1002      '  nome inglese del paese
 Const LOCALE_SDECIMAL = &HE                   '  separatore decimale
 Const LOCALE_SDAYNAME7 = &H30        '  nome lungo per domenica
 Const LOCALE_SDAYNAME6 = &H2F        '  nome lungo per sabato
 Const LOCALE_SDAYNAME5 = &H2E        '  nome lungo per venerdì
 Const LOCALE_SDAYNAME4 = &H2D        '  nome lungo per giovedì
 Const LOCALE_SDAYNAME3 = &H2C        '  nome lungo per mercoledì
 Const LOCALE_SDAYNAME2 = &H2B        '  nome lungo per martedì
 Const LOCALE_SDAYNAME1 = &H2A        '  nome lungo per lunedì
 Const LOCALE_SDATE = &H1D                 '  separatore di data
 Const LOCALE_SCURRENCY = &H14        '  simbolo locale valuta
 Const LOCALE_SCOUNTRY = &H6            '  nome localizzato del paese
 Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  nome abbreviato per settembre
 Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  nome abbreviato per agosto
 Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  nome abbreviato per luglio
 Const LOCALE_SABBREVMONTHNAME6 = &H49        '  nome abbreviato per giugno
 Const LOCALE_SABBREVMONTHNAME5 = &H48        '  nome abbreviato per maggio
 Const LOCALE_SABBREVMONTHNAME4 = &H47        '  nome abbreviato per aprile
 Const LOCALE_SABBREVMONTHNAME3 = &H46        '  nome abbreviato per marzo
 Const LOCALE_SABBREVMONTHNAME2 = &H45        '  nome abbreviato per febbraio
 Const LOCALE_SABBREVMONTHNAME13 = &H100F
 Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  nome abbreviato per dicembre
 Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  nome abbreviato per novembre
 Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  nome abbreviato per ottobre
 Const LOCALE_SABBREVMONTHNAME1 = &H44        '  nome abbreviato per gennaio
 Const LOCALE_SABBREVLANGNAME = &H3         '  nome della lingua abbreviato
 Const LOCALE_SABBREVDAYNAME7 = &H37        '  nome abbreviato per domenica
 Const LOCALE_SABBREVDAYNAME6 = &H36        '  nome abbreviato per sabato
 Const LOCALE_SABBREVDAYNAME5 = &H35        '  nome abbreviato per venerdì
 Const LOCALE_SABBREVDAYNAME4 = &H34        '  nome abbreviato per giovedì
 Const LOCALE_SABBREVDAYNAME3 = &H33        '  nome abbreviato per mercoledì
 Const LOCALE_SABBREVDAYNAME2 = &H32        '  nome abbreviato per martedì
 Const LOCALE_SABBREVDAYNAME1 = &H31        '  nome abbreviato per lunedì
 Const LOCALE_SABBREVCTRYNAME = &H7         '  nome abbreviato del paese
 Const LOCALE_S2359 = &H29                                '  identificatore PM
 Const LOCALE_S1159 = &H28                                 '  identificatore AM
 Const LOCALE_NOUSEROVERRIDE = &H80000000  '  non utilizza impostazioni utente
 Const LOCALE_ITLZERO = &H25        '  zero iniziali nei campi di ora
 Const LOCALE_IPOSSYMPRECEDES = &H54        '  simb valuta precede importo positivo
 Const LOCALE_IPOSSIGNPOSN = &H52        '  posizione segno positivo
 Const LOCALE_IPOSSEPBYSPACE = &H55        '  simb valuta sep. da spazio da imp. positivo
 Const LOCALE_INEGSYMPRECEDES = &H56        '  simb valuta precede importo negativo
 Const LOCALE_INEGSIGNPOSN = &H53        '  posizione segno negativo
 Const LOCALE_INEGSEPBYSPACE = &H57        '  simb valuta sep. da spazio da imp. negativo
 Const LOCALE_INEGCURR = &H1C        '  modalità valuta negativa
 Const LOCALE_IMONLZERO = &H27        '  zero iniziali nei campi del mese
 Const LOCALE_IMEASURE = &HD         '  0 = metrico, 1 = US
 Const LOCALE_ILZERO = &H12        '  zero iniziali per i decimali
 Const LOCALE_ILDATE = &H22        '  ordinamento formato data lungo
 Const LOCALE_ILANGUAGE = &H1         '  ID lingua
 Const LOCALE_IINTLCURRDIGITS = &H1A        '  # cifre valuta internazionale
 Const LOCALE_IDIGITS = &H11        '  numero di cifre decimali
 Const LOCALE_IDEFAULTLANGUAGE = &H9         '  ID di lingua predefinito
 Const LOCALE_IDEFAULTCOUNTRY = &HA         '  codice di paese predefinito
 Const LOCALE_IDEFAULTCODEPAGE = &HB         '  tabella codici predefinita
 Const LOCALE_IDAYLZERO = &H26        '  zero iniziali nei campi del giorno
 Const LOCALE_IDATE = &H21        '  ordinamento formato data breve
 Const LOCALE_ICURRENCY = &H1B        '  modalità valuta positiva
 Const LOCALE_ICURRDIGITS = &H19        '  # cifre valuta locale
 Const LOCALE_ICOUNTRY = &H5         '  codice di paese
 Const LOCALE_ICENTURY = &H24        '  indicatore formato secolo
 
 'utente locale
 
 Const LOCALE_USER_DEFAULT& = &H400

Public Const DATE_LONGDATE As Long = &H2
Public Const DATE_SHORTDATE As Long = &H1

Public Const HWND_BROADCAST As Long = &HFFFF&
Public Const WM_SETTINGCHANGE As Long = &H1A

Public Declare Function PostMessage Lib "user32" _
   Alias "PostMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long

Public Declare Function EnumDateFormats Lib "kernel32" _
   Alias "EnumDateFormatsA" _
  (ByVal lpDateFmtEnumProc As Long, _
   ByVal Locale As Long, _
   ByVal dwFlags As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long

Public Declare Function SetLocaleInfo Lib "kernel32" _
    Alias "SetLocaleInfoA" _
   (ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String) As Long

' acquisisce le informazioni locali

Function InfoLocali(LCType As Long) As String
  Dim lngret As Long
  Dim Locale As Long
  Dim lpLCData As String
  Dim cchData As Long

  Locale = LOCALE_USER_DEFAULT&
  lpLCData = String$(255, 0)                         ' stringa contenente 255 valori "0"
  cchData = Len(lpLCData)
  lngret = GetLocaleInfo(Locale, LCType, lpLCData, cchData)
  If cchData = 0 Then
    lpLCData = String$(lngret, 0)
    cchData = Len(lpLCData)
    lngret = GetLocaleInfo(Locale, LCType, lpLCData, cchData)
  End If
  InfoLocali = Left(lpLCData, lngret - 1)
End Function

'setta le informazion di localizzazione del computer

Sub SetInfoLocali(LCType As Long, ByVal valore As String)
  Dim lngret As Long
  Dim Locale As Long
  Dim lpLCData As String
  Dim cchData As Long

  Locale = LOCALE_USER_DEFAULT&
  lpLCData = valore
  lpLCData = lpLCData & String$(255 - Len(valore), 0)                       ' stringa contenente 255 valori "0"
  cchData = Len(lpLCData)
  lngret = SetLocaleInfo(Locale, LCType, lpLCData)
  
  'InfoLocali = Left(lpLCData, lngret - 1)
End Sub


Public Function fGetUserLocaleInfo(ByVal lLocaleID As Long, _
            ByVal lLCType As Long) As String

Dim sReturn As String
Dim lReturn As Long
'
' acquisisce la lunghezza del formato desiderato
'
'
lReturn = GetLocaleInfo(lLocaleID, lLCType, sReturn, Len(sReturn))
'
' se ottengo la lunghezza del buffer richiama la funzione
'
If lReturn Then
    '
    ' introduce spazi nel buffer
    '
    sReturn = Space$(lReturn)
    '
    ' acquisisce la data
    '
    lReturn = GetLocaleInfo(lLocaleID, lLCType, sReturn, Len(sReturn))
    '
    ' se Ok rimuove i null
    '
    If lReturn Then
        fGetUserLocaleInfo = Left$(sReturn, lReturn - 1)
    End If
End If
    
End Function


Private Function fStringFromPointer(sString As Long) As String
Dim lPos    As Long
Dim sBuffer As String
'
'
'
sBuffer = Space$(128)
'
' copia il puntatore stringa nel valore di ritorno
' della funzione fEnumDates
'
Call CopyMemory(ByVal sBuffer, sString, ByVal Len(sBuffer))
'
' rimuove i Null
'
lPos = InStr(sBuffer, Chr$(0))

If lPos Then
    fStringFromPointer = Left$(sBuffer, lPos - 1)
End If
End Function

