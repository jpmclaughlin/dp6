VERSION 5.00
Begin VB.UserControl ControlloPacco 
   Alignable       =   -1  'True
   BackColor       =   &H0083FFFF&
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7935
   ControlContainer=   -1  'True
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   HasDC           =   0   'False
   MaskColor       =   &H00FFC0FF&
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   ToolboxBitmap   =   "ControlloPacco.ctx":0000
   Begin VB.Label LabelSpessore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Spessore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label LabelAltezza 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Altezz."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1620
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LabelBase 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "ControlloPacco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ============================================================
' BundleDraw6 per disegno pacco esagono e quadro completo di
' 1) Disegno quote
' 2) Selezione Pollici o mm
' 3) Disegno spessore proporzionale

' Dati di ingresso:

' aConfig = Selezioni varie
'           bit 0 : unità di misura : 0 = mm   1 = inch
'           bit 2 : tipo di disegno : 0 = pacco 1=tubo
'           bit 3 : disegna spessore: 0 = no 1=si
'           bit 4 : disegna label   : 0 = no 1=si
'           bit 5 : ricalcolo profili: 0 = no  1=si
' aTube_Width = larghezza tubo in   mm x 10   o inch * 100
' aTube_Height = altezza tubo in    mm x 10   o inch * 100
' aTube_Tickness = spessore tubo in mm x 100  o inch * 1000
' aCounted  = Tubi presenti nel pacco
' Row_01 .... Row_50 = Numero tubi per ogni fila
' ============================================================

Private m_Sfondo As Long

'Variabili proprietà:
Private mCounted As Long

Public aTipoPacco As Long
Public aTipoTubo As Long
Public aConfig As Integer
Public aTube_Width As Long
Public aTube_Height As Long
Public aTube_Tickness As Long
Public aTipoProfilo As Integer

Private Const MAX_FILE As Integer = 50

'Costanti per disegno quotature (Pixel)
Private Const LarghezzaFreccia As Long = 4
Private Const LunghezzaFreccia As Long = 10
Private Const SpazioQuotaLato As Long = 50
Private Const SpazioQuotaSotto As Long = 40

'dati per disegno quote  (pixel)
Private QuoteXMin As Long
Private QuoteXMax As Long
Private QuoteYMin As Long
Private QuoteYMax As Long
Private QuoteXBaseMin As Long
Private QuoteXBaseMax As Long
Private QuoteYLatoMax As Long
Private PosizioneYFilaMax As Long
'
Private AltezzaTuboPixel As Long        'altezza tubo
Private LarghezzaTuboPixel As Long      'larghezza tubo
Private SpessoreTuboPixel As Long       'spessore tubo
Private PassoTuboPixel As Long
Private PassoFilaPixel As Long

'Composizione pacco
Private Row(MAX_FILE) As Long
Private NumeroFile As Long
Private TubiFilaMax As Long
Private PosizioneFilaMax As Long
Private mInch As Boolean
Private mTubeRound As Boolean
Private mBundleHex As Boolean
Private mDisegnaTubo As Boolean
Private mDisegnaSpessore As Boolean
Private mDisegnaLabel As Boolean
Private mProfili As Boolean

' dati calcolati
Private TubiPacco As Long
Private LatoPaccoHex As Long
Private BasePaccoHex As Long
Private AltezzaPacco As Long
Private LarghezzaPacco As Long

Public InDisegno As Boolean

'evento

Public Event InizioDisegno()
Public Event Disegnato()

'=============================================================================================
'                     DICHIARAZIONE DELLE PROPRIETA'
'=============================================================================================

Public Property Let aCounted(ByVal New_Counted As Long)
    If mDisegnaTubo Then
        mCounted = 0
    Else
        mCounted = New_Counted
    End If
    Let_s_draw      'Ridisegna
End Property

Public Property Let TubiFila(Index As Integer, Value As Integer)
    If Index >= 1 And Index <= MAX_ROWS Then
        Row(Index) = Value
    End If
End Property

'=============================================================================================
'                            METODI PUBBLICI
'=============================================================================================
Public Sub Refresh()

    ' crea un evento prima di iniziare a disegnare
    InDisegno = True
    RaiseEvent InizioDisegno
    
    Cls
    Call Let_s_draw
    
    ' quando ha finito il disegno genera l'evento disegnato
    
    InDisegno = False
    RaiseEvent Disegnato
    
End Sub
'=============================================================================================
'=============================================================================================

Private Sub UserControl_Initialize()
    aConfig = 12 ' mm,pacco
    aTube_Width = 600
    aTube_Height = 600
    aTube_Tickness = 10
    aTipoTubo = 1
    aTipoPacco = 1
    Row(1) = 2
    Row(2) = 3
    Row(3) = 2
    aCounted = 1
    InDisegno = False
End Sub

Private Sub UserControl_Paint()
    Let_s_draw
End Sub

Private Sub LeggiParametri()
    Dim i As Integer
    
    If (aConfig And &H1&) Then
        mInch = False
    Else
        mInch = True
    End If
    If (aConfig And &H2&) Then
        mDisegnaTubo = True
    Else
        mDisegnaTubo = False
    End If
    If (aConfig And &H4&) Then
        mDisegnaSpessore = True
    Else
        mDisegnaSpessore = False
    End If
    If (aConfig And &H8&) Then
        mDisegnaLabel = True
    Else
        mDisegnaLabel = False
    End If
    ' controlla se ci sono profili
    
       mProfili = (aConfig And &H10&)
   
    ' controlla il tipo di pacco
    If aTipoPacco = 1 And Not mDisegnaTubo Then
        mBundleHex = True
    Else
        mBundleHex = False
    End If
    ' controlla il tipo di tubo
    If aTipoTubo = 1 Then
        mTubeRound = True
    Else
        mTubeRound = False
    End If
    ' disegna il tubo
    If mDisegnaTubo Then
        Row(1) = 1
        For i = 2 To MAX_ROWS
            Row(i) = 0
        Next
        mCounted = 0
    End If

    ' calcolo numero file
    NumeroFile = MAX_FILE
    For i = MAX_FILE To 1 Step -1
        If Row(i) <= 0 Then NumeroFile = i - 1
    Next i
    ' calcola fila max
    TubiFilaMax = 0
    For i = 1 To NumeroFile
        If TubiFilaMax < Row(i) Then
            TubiFilaMax = Row(i)
            PosizioneFilaMax = i
        End If
    Next i
    ' per evitare successive divisioni per 0
    If TubiFilaMax < 1 Then TubiFilaMax = 1

End Sub

Private Sub CalcolaDimensioniTuboInPixel()
    Dim PixelDisponibili As Integer
    
    ' LARGHEZZA
    LarghezzaTuboPixel = 0
    PixelDisponibili = ScaleWidth - SpazioQuotaLato
    If PixelDisponibili < 0 Then PixelDisponibili = 0
    If TubiFilaMax > 0 Then
        LarghezzaTuboPixel = PixelDisponibili / TubiFilaMax
        If (LarghezzaTuboPixel Mod 2) > 0 Then LarghezzaTuboPixel = LarghezzaTuboPixel - 1
        If (LarghezzaTuboPixel * TubiFilaMax) > PixelDisponibili Then
            LarghezzaTuboPixel = LarghezzaTuboPixel - 2
        End If
    End If
    If LarghezzaTuboPixel < 2 Then LarghezzaTuboPixel = 2
    
    ' ALTEZZA
    If aTube_Height = aTube_Width Then
        AltezzaTuboPixel = LarghezzaTuboPixel
    Else
        If aTube_Width > 0 Then
            AltezzaTuboPixel = (LarghezzaTuboPixel * aTube_Height) / aTube_Width
            If (AltezzaTuboPixel Mod 2) > 0 Then AltezzaTuboPixel = AltezzaTuboPixel - 1
        Else
            AltezzaTuboPixel = LarghezzaTuboPixel
        End If
    End If
    If AltezzaTuboPixel < 2 Then AltezzaTuboPixel = 2
    
    ' SPESSORE
    SpessoreTuboPixel = 0
    If aTube_Width > 0 Then
        SpessoreTuboPixel = (LarghezzaTuboPixel * aTube_Tickness) / (aTube_Width * 10&)  ' Thickness è sempre in una unità 10 volte più piccola di Width
    End If
    If SpessoreTuboPixel >= (AltezzaTuboPixel / 2) Then SpessoreTuboPixel = (AltezzaTuboPixel / 2) - 1
    If SpessoreTuboPixel >= (LarghezzaTuboPixel / 2) Then SpessoreTuboPixel = (LarghezzaTuboPixel / 2) - 1
    If SpessoreTuboPixel < 1 Then SpessoreTuboPixel = 1
    
    ' PASSO TUBO
    PassoTuboPixel = LarghezzaTuboPixel
    
    ' PASSO FILA
    If mTubeRound Then
        PassoFilaPixel = AltezzaTuboPixel * 866& / 1000& + 1
    Else
        PassoFilaPixel = AltezzaTuboPixel
    End If
    
End Sub

Private Sub Let_s_draw()
    Call LeggiParametri
    Call CalcolaDimensioniTuboInPixel
    Call DisegnaTubi
    If mDisegnaTubo Then
        If mTubeRound Then
            DisegnaQuotaLarghezza
            LabelAltezza.Visible = False
        Else
            DisegnaQuotaLarghezza
            DisegnaQuotaAltezza
        End If
        If mDisegnaSpessore Then
            DisegnaQuotaSpessore
        Else
            LabelSpessore.Visible = False
        End If
    Else
        LabelSpessore.Visible = False
        If mBundleHex Then
            DisegnaQuoteLatoHex
        Else
            DisegnaQuotaLarghezza
            DisegnaQuotaAltezza
        End If
    End If
End Sub

Private Sub DisegnaTubi()
    Dim MezzaLarghezzaTubo As Long
    Dim MezzaAltezzaTubo As Long
    Dim Color As Long
    Dim cx As Single
    Dim cy As Single
    Dim ContaFile As Long
    Dim ContaTubiRow As Long
    Dim ContaTubiPacco As Long
    Dim RaggioOrizz As Single
    Dim RaggioVert As Single
    
    ScaleMode = 3               'dimensioni grafiche in pixel
    Color = &H0                 'colore linea
    FillColor = &H606060        'colore riempimento

    MezzaAltezzaTubo = AltezzaTuboPixel / 2
    MezzaLarghezzaTubo = LarghezzaTuboPixel / 2
    DrawWidth = SpessoreTuboPixel
    
    QuoteXMin = ScaleWidth
    QuoteXMax = 0
    QuoteXBaseMin = ScaleWidth
    QuoteXBaseMax = 0
    QuoteYMin = ScaleHeight
    QuoteYMax = 0
    ' calcola coordinata y della prima fila
    cy = ScaleHeight - SpazioQuotaSotto - MezzaAltezzaTubo
    ' calcola raggio tubi tondi
    If (SpessoreTuboPixel Mod 2) > 0 Then
        RaggioOrizz = MezzaLarghezzaTubo - (SpessoreTuboPixel / 2) + 0.5
        RaggioVert = MezzaAltezzaTubo - (SpessoreTuboPixel / 2) + 0.5
    Else
        RaggioOrizz = MezzaLarghezzaTubo - (SpessoreTuboPixel / 2)
        RaggioVert = MezzaAltezzaTubo - (SpessoreTuboPixel / 2)
    End If
    
    For ContaFile = 1 To MAX_FILE
        ' calcola coordinata x del primo tubo della fila
        cx = ((ScaleWidth - SpazioQuotaLato) - ((Row(ContaFile) - 1) * PassoTuboPixel)) / 2
        
        For ContaTubiRow = 1 To Row(ContaFile)
            ' memorizza coordinate limite dei tubi per quotature
            If mBundleHex Then
                If QuoteXMin > (cx - MezzaLarghezzaTubo * 1155& / 1000&) Then QuoteXMin = (cx - MezzaLarghezzaTubo * 1155& / 1000&)
                If QuoteXMax < (cx + MezzaLarghezzaTubo * 1155& / 1000&) Then
                    QuoteXMax = (cx + MezzaLarghezzaTubo * 1155& / 1000&)
                    PosizioneYFilaMax = cy
                End If
                If ContaFile = 1 Then
                    If QuoteXBaseMin >= (cx - MezzaLarghezzaTubo * 577& / 1000&) Then QuoteXBaseMin = (cx - MezzaLarghezzaTubo * 577& / 1000&)
                    If QuoteXBaseMax <= (cx + MezzaLarghezzaTubo * 577& / 1000&) Then QuoteXBaseMax = (cx + MezzaLarghezzaTubo * 577& / 1000&)
                End If
            Else
                If QuoteXMin >= (cx - MezzaLarghezzaTubo) Then QuoteXMin = (cx - MezzaLarghezzaTubo - 1)
                If QuoteXMax <= (cx + MezzaLarghezzaTubo) Then QuoteXMax = (cx + MezzaLarghezzaTubo + 1)
            End If
            If QuoteXMin >= (cx - MezzaLarghezzaTubo) Then QuoteXMin = (cx - MezzaLarghezzaTubo - 1)
            If QuoteXMax <= (cx + MezzaLarghezzaTubo) Then QuoteXMax = (cx + MezzaLarghezzaTubo + 1)
            If QuoteYMin >= (cy - MezzaAltezzaTubo) Then QuoteYMin = (cy - MezzaAltezzaTubo - 1)
            If QuoteYMax <= (cy + MezzaAltezzaTubo) Then QuoteYMax = (cy + MezzaAltezzaTubo + 1)
            ' conta i tubi
            ContaTubiPacco = ContaTubiPacco + 1
            If ContaTubiPacco > mCounted Then
                If mDisegnaTubo Then        'colore linea
                    Color = &H0
                Else
                    Color = &H808080
                End If
                FillColor = &HE0E0E0
            End If
            ' disegna il tubo
            If mTubeRound Then
                Circle (cx, cy), RaggioOrizz, Color
            Else
                Line (cx - RaggioOrizz, cy - RaggioVert)-Step(RaggioOrizz * 2, RaggioVert * 2), Color, B
            End If
            ' calcola coordinata x del prossimo tubo della fila
            cx = cx + PassoTuboPixel
        Next ContaTubiRow
        ' calcola coordinata y della prossima fila
        If (ContaFile < MAX_FILE) Then
            If (mTubeRound) And (Row(ContaFile) = Row(ContaFile + 1)) Then
                cy = cy - AltezzaTuboPixel
            Else
                cy = cy - PassoFilaPixel
            End If
        Else
            cy = cy - PassoFilaPixel
        End If
    Next ContaFile
        
End Sub


Private Sub DisegnaQuotaLarghezza()
    Dim i As Integer
    Dim QBase_Y As Integer
    Dim QAltezza_X As Integer

    DrawWidth = 1
    
    'calcolo larghezza pacco:
    If mProfili Then
       Select Case aTipoProfilo
       Case 1
           LarghezzaPacco = aTube_Width * TubiFilaMax
       Case 2
           LarghezzaPacco = aTube_Width * (TubiFilaMax + 2)
       End Select
    Else
       LarghezzaPacco = aTube_Width * TubiFilaMax
    End If
    '--------------------
    
    'disegno QUOTA LARGHEZZA
    DrawWidth = 1
    QBase_Y = (ScaleHeight - LabelBase.Height / 2 - 2)
    DrawStyle = vbDot
    'linee laterali quota
    QuoteXMin = QuoteXMin + 1
    QuoteXMax = QuoteXMax - 1
    If mDisegnaTubo And mTubeRound Then
        Line (QuoteXMin, ScaleHeight - SpazioQuotaSotto - AltezzaTuboPixel / 2)-(QuoteXMin, QBase_Y), 0
        Line (QuoteXMax, ScaleHeight - SpazioQuotaSotto - AltezzaTuboPixel / 2)-(QuoteXMax, QBase_Y), 0
    Else
        Line (QuoteXMin, ScaleHeight - SpazioQuotaSotto)-(QuoteXMin, QBase_Y), 0
        Line (QuoteXMax, ScaleHeight - SpazioQuotaSotto)-(QuoteXMax, QBase_Y), 0
    End If
    DrawStyle = vbSolid
    'linea quota
    Line (QuoteXMin, QBase_Y)-(QuoteXMax, QBase_Y), 0
    'freccia sinistra
    Line (QuoteXMin, QBase_Y)-Step(LunghezzaFreccia, LarghezzaFreccia), 0
    Line (QuoteXMin, QBase_Y)-Step(LunghezzaFreccia, -LarghezzaFreccia), 0
    'freccia destra
    Line (QuoteXMax, QBase_Y)-Step(-LunghezzaFreccia, LarghezzaFreccia), 0
    Line (QuoteXMax, QBase_Y)-Step(-LunghezzaFreccia, -LarghezzaFreccia), 0
    'label quota
    LabelBase.Visible = False
    If mDisegnaLabel Then
        If mInch Then
            If mDisegnaTubo Then
                LabelBase.Caption = Format(LarghezzaPacco / 100#, "0.00")
            Else
                LabelBase.Caption = Format(LarghezzaPacco / 100#, "0.0")
            End If
        Else
            If mDisegnaTubo Then
                LabelBase.Caption = Format(LarghezzaPacco / 100#, "0.00")
            Else
                LabelBase.Caption = Format(LarghezzaPacco / 100#, "0.0")
            End If
        End If
        LabelBase.Top = QBase_Y - LabelBase.Height / 2
        LabelBase.Left = (ScaleWidth - SpazioQuotaLato - LabelBase.Width) / 2
        LabelBase.Visible = True
    End If
End Sub


Private Sub DisegnaQuotaAltezza()
    Dim i As Integer
    Dim QBase_Y As Integer
    Dim QAltezza_X As Integer

    DrawWidth = 1
 
    'calcolo altezza pacco:
    AltezzaPacco = 0#
    If mTubeRound Then
        For i = 1 To NumeroFile
            If i > 1 And Row(i) <> Row(i - 1) Then
                AltezzaPacco = AltezzaPacco + aTube_Height * 0.866
            Else
                AltezzaPacco = AltezzaPacco + aTube_Height
            End If
        Next i
    Else 'pacco quadro
        If mProfili Then
           Select Case aTipoProfilo
           Case 1
               AltezzaPacco = aTube_Height * NumeroFile
           Case 2
               AltezzaPacco = (aTube_Height + (2 * aTube_Tickness / 10) * (1 + 2 * Abs(TubiFilaMax > 2))) * NumeroFile
           End Select
        Else
           AltezzaPacco = aTube_Height * NumeroFile
        End If
    End If
    
    'disegno QUOTA ALTEZZA
    QAltezza_X = ScaleWidth - SpazioQuotaLato / 2
    DrawStyle = vbDot
    'linee laterali quota
    Line (QuoteXMax, QuoteYMax)-(QAltezza_X, QuoteYMax), 0
    Line (QuoteXMax, QuoteYMin)-(QAltezza_X, QuoteYMin), 0
    'linea quota
    DrawStyle = vbSolid
    Line (QAltezza_X, QuoteYMax)-(QAltezza_X, QuoteYMin), 0
    'freccia sopra
    Line (QAltezza_X, QuoteYMax)-Step(-LarghezzaFreccia, -LunghezzaFreccia), 0
    Line (QAltezza_X, QuoteYMax)-Step(LarghezzaFreccia, -LunghezzaFreccia), 0
    'freccia sotto
    Line (QAltezza_X, QuoteYMin)-Step(LarghezzaFreccia, LunghezzaFreccia), 0
    Line (QAltezza_X, QuoteYMin)-Step(-LarghezzaFreccia, LunghezzaFreccia), 0
    'label quota
    LabelAltezza.Visible = False
    If mDisegnaLabel Then
        If mInch Then
            If mDisegnaTubo Then
                LabelAltezza.Caption = Format(AltezzaPacco / 100#, "0.00")
            Else
                LabelAltezza.Caption = Format(AltezzaPacco / 100#, "0.0")
            End If
        Else
            If mDisegnaTubo Then
                LabelAltezza.Caption = Format(AltezzaPacco / 100#, "0.00")
            Else
                LabelAltezza.Caption = Format(AltezzaPacco / 100#, "0.0")
            End If
        End If
        LabelAltezza.Top = QuoteYMin + (QuoteYMax - QuoteYMin) / 2 - LabelAltezza.Height / 2
        LabelAltezza.Left = QAltezza_X - LabelAltezza.Width / 2
        LabelAltezza.Visible = True
    End If
End Sub

Private Sub DisegnaQuoteLatoHex()
    Dim X1 As Single
    Dim Y1 As Single
    Dim X2 As Single
    Dim Y2 As Single
    Dim TubiLatoHex As Integer
    Dim QBase_Y As Integer
    Dim OffsetLato As Integer
    
    TubiLatoHex = PosizioneFilaMax 'numero tubi per calcolo quota lato

    'calcolo quota latohex :
    LatoPaccoHex = ((aTube_Height * ((TubiLatoHex - 2#) + 1#)) + 2# * (aTube_Height / (2# * Sqr(3)))) / 10#
    
    'calcolo quota base hex:
    BasePaccoHex = ((aTube_Height * ((Row(1) - 2#) + 1#)) + 2# * (aTube_Height / (2# * Sqr(3)))) / 10#
    
    'disegno QUOTA BASE HEX
    DrawWidth = 1
    QBase_Y = (ScaleHeight - LabelBase.Height / 2 - 2)
    DrawStyle = vbDot
    'linee laterali quota
    Line (QuoteXBaseMin, ScaleHeight - SpazioQuotaSotto)-(QuoteXBaseMin, QBase_Y), 0
    Line (QuoteXBaseMax, ScaleHeight - SpazioQuotaSotto)-(QuoteXBaseMax, QBase_Y), 0
    DrawStyle = vbSolid
    'linea quota
    Line (QuoteXBaseMin, QBase_Y)-(QuoteXBaseMax, QBase_Y), 0
    'freccia sinistra
    Line (QuoteXBaseMin, QBase_Y)-Step(LunghezzaFreccia, LarghezzaFreccia), 0
    Line (QuoteXBaseMin, QBase_Y)-Step(LunghezzaFreccia, -LarghezzaFreccia), 0
    'freccia destra
    Line (QuoteXBaseMax, QBase_Y)-Step(-LunghezzaFreccia, LarghezzaFreccia), 0
    Line (QuoteXBaseMax, QBase_Y)-Step(-LunghezzaFreccia, -LarghezzaFreccia), 0
    'label quota
        LabelBase.Visible = False
    If mDisegnaLabel Then
        If mInch Then
            LabelBase.Caption = Format(BasePaccoHex / 10#, "0.0")
        Else
            LabelBase.Caption = Format(BasePaccoHex / 10#, "0.0")
        End If
        LabelBase.Top = QBase_Y - LabelBase.Height / 2
        LabelBase.Left = (QuoteXBaseMax + QuoteXBaseMin) / 2 - LabelBase.Width / 2
        LabelBase.Visible = True
    End If
    'disegno QUOTA LATO HEX
    X1 = QuoteXMax
    Y1 = PosizioneYFilaMax
    X2 = QuoteXBaseMax
    Y2 = ScaleHeight - SpazioQuotaSotto
    DrawWidth = 1
    DrawStyle = vbDot
    OffsetLato = 20
    'linee laterali quota
    Line (X1, Y1)-(X1 + SpazioQuotaLato - OffsetLato, Y1 + (SpazioQuotaLato - OffsetLato) * 577& / 1000&), 0
    Line (X2, Y2)-(X2 + SpazioQuotaLato - OffsetLato, Y2 + (SpazioQuotaLato - OffsetLato) * 577& / 1000&), 0
    'linea quota
    DrawStyle = vbSolid
    X1 = X1 + SpazioQuotaLato - OffsetLato - LarghezzaFreccia
    Y1 = Y1 + (SpazioQuotaLato - OffsetLato - LarghezzaFreccia) * 577& / 1000&
    X2 = X2 + SpazioQuotaLato - OffsetLato - LarghezzaFreccia
    Y2 = Y2 + (SpazioQuotaLato - OffsetLato - LarghezzaFreccia) * 577& / 1000&
    Line (X1, Y1)-(X2, Y2), 0
    'freccia sopra
    Line (X1 - 1, Y1 + 1)-Step(0, LunghezzaFreccia), 0
    Line (X1 - 1, Y1 + 1)-Step(-LunghezzaFreccia, LarghezzaFreccia), 0
    'freccia sotto
    Line (X2, Y2)-Step(0, -LunghezzaFreccia), 0
    Line (X2, Y2)-Step(LunghezzaFreccia, -LarghezzaFreccia), 0
    'label quota
    LabelAltezza.Visible = False
    If mDisegnaLabel Then
        If mInch Then
            LabelAltezza.Caption = Format(LatoPaccoHex / 10#, "0.0")
        Else
            LabelAltezza.Caption = Format(LatoPaccoHex / 10#, "0.0")
        End If
        LabelAltezza.Top = (Y1 + Y2) / 2# - LabelAltezza.Height / 2
        LabelAltezza.Left = (X1 + X2) / 2# - LabelAltezza.Width / 2
        LabelAltezza.Visible = True
    End If
End Sub


Private Sub DisegnaQuotaSpessore()
    Dim i As Integer
    Dim QCentroX As Integer
    DrawWidth = 1
    DrawStyle = vbSolid

    'disegno QUOTA SPESSORE
    QCentroX = (ScaleWidth - SpazioQuotaLato) / 2
    If mDisegnaLabel Then
        'linea quota verticale
        Line (QCentroX, QuoteYMin)-Step(0, -LunghezzaFreccia * 2), 0
        'linea quota orizzontale
        Line Step(0, 0)-Step(LunghezzaFreccia * 2, 0), 0
    Else
        'linea quota verticale
        Line (QCentroX, 0)-(QCentroX, QuoteYMin), 0
        'linea quota orizzontale
        Line (QCentroX, 0)-Step(LunghezzaFreccia * 4, 0), 0
    End If
    'freccia sopra
    Line (QCentroX, QuoteYMin)-Step(-LarghezzaFreccia, -LunghezzaFreccia), 0
    Line (QCentroX, QuoteYMin)-Step(LarghezzaFreccia, -LunghezzaFreccia), 0
    'freccia sotto
    Line (QCentroX, QuoteYMin + SpessoreTuboPixel)-Step(0, LunghezzaFreccia * 2), 0
    Line (QCentroX, QuoteYMin + SpessoreTuboPixel)-Step(-LarghezzaFreccia, LunghezzaFreccia), 0
    Line (QCentroX, QuoteYMin + SpessoreTuboPixel)-Step(LarghezzaFreccia, LunghezzaFreccia), 0
'    'label quota
    LabelSpessore.Visible = False
    If mDisegnaLabel Then
        If mInch Then
            LabelSpessore.Caption = Format(aTube_Tickness / 1000#, "0.00")
        Else
            LabelSpessore.Caption = Format(aTube_Tickness / 1000#, "0.00")
        End If
        LabelSpessore.Top = QuoteYMin - LunghezzaFreccia * 2 - LabelSpessore.Height / 2
        LabelSpessore.Left = QCentroX + LunghezzaFreccia * 2 + 10
        LabelSpessore.Visible = True
    End If
End Sub

Property Let ColoreSfondo(ByVal Colore As Long)
m_Sfondo = Colore
UserControl.BackColor = Colore
LabelSpessore.BackColor = Colore
LabelAltezza.BackColor = Colore
LabelBase.BackColor = Colore
End Property
