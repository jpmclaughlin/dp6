VERSION 5.00
Begin VB.UserControl Barra 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   PropertyPages   =   "BarraOpzioni.ctx":0000
   ScaleHeight     =   1290
   ScaleWidth      =   15360
   ToolboxBitmap   =   "BarraOpzioni.ctx":0010
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   0
      Left            =   14460
      Top             =   30
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   12
      Left            =   13200
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   11
      Left            =   12000
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   10
      Left            =   10800
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   9
      Left            =   9600
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   8
      Left            =   8400
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   7
      Left            =   7200
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   6
      Left            =   6000
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   5
      Left            =   4800
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   4
      Left            =   3600
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   3
      Left            =   2400
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   2
      Left            =   1200
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Pulsante 
      Height          =   1095
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Barra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'NOTA: Al fine di sostituire la barra opzioni costruita con pulsanti option
'      nel DP6 senza cambiare la struttura del programma, la proprietà
'      "selezionato" e l'evento click restituiscono il valore dell'indice+1

'CONST tipi pulsanti presenti nella option bar

Enum IndicePuls
   
     pOrdini = 1
     pstorico = 2
     pMappa = 3
     pPacco = 4
     pRegge = 5
     pPesa = 6
     pAllarmi = 7
     pEntrata = 8
     pSmussatrice = 9
     pFilettatura = 10
     pWalkingBeam = 11
     pTaglio = 12
End Enum

' VARIABILI locali
Private m_Selezionato As IndicePuls
Private PrimoCiclo As Boolean
Private m_PulsanteID As IndicePuls
Private m_Allarme As Boolean
Public Bloccata As Boolean
' EVENTO dichiarazione: generato dal click su un pulsante della barra
Public Event RipetizioneTasto()
Public Event PulsantePremuto(ByVal Index As IndicePuls)

'EVENTO: click su controllo immagine (pulsanti); restituisce l'indice del pulsante
'la visibilità è pubblica: metodo "Pulsante_click x"

Public Sub Pulsante_Click(Index As Integer)

If Pulsante(Index).Enabled = True And Not (Bloccata) Then
   If Index <> m_Selezionato Or PrimoCiclo = False Then
      Pulsante(m_Selezionato) = LoadResPicture(100 + m_Selezionato, 0)
      Pulsante(Index) = LoadResPicture(200 + Index, 0)
      m_Selezionato = Index
      PrimoCiclo = True
      RaiseEvent PulsantePremuto(Index)
   Else
      RaiseEvent RipetizioneTasto
   End If
End If
End Sub

Private Sub Timer1_Timer()
Static All As Boolean

If Allarme = True Then
   All = Not (All)
   If All = False Then
       If m_Selezionato <> pAllarmi Then
          Pulsante(7) = LoadResPicture(107, 0)
       Else
          Pulsante(7) = LoadResPicture(207, 0)
       End If
   Else
       Pulsante(7) = LoadResPicture(213, 0)
   End If
End If
End Sub

' INIZIALIZZA: valori e immagini pulsanti

Private Sub UserControl_Initialize()
Dim i As Byte

For i = 1 To 12
   Pulsante(i) = LoadResPicture(100 + i, 0)
   Pulsante(i).Top = 0
Next

PrimoCiclo = False
m_Selezionato = 1
m_Allarme = False
Timer1.Enabled = False
End Sub
'Metodo: scrive la posizione (0,x) del pulsante specificato rispetto allo 0,0 del controllo
Sub Posizione(ByVal Puls As IndicePuls, ByVal Left)
Pulsante(Puls).Left = Left
End Sub
'PROP: sola lettura; identifica l'indice del pulsante attualmente selezionato
Property Get Selezionato() As Byte
Selezionato = m_Selezionato
End Property
'PROP: L/S; (dopo aver assegnato un indirizzo a pulsanteID) assegna la posizione
Property Let PulsantePos(ByVal Left As Long)
Pulsante(m_PulsanteID).Left = Left
PropertyChanged "PulsantePos"
UserControl.Refresh
End Property
'PROP: L/S; (dopo aver assegnato un indirizzo a pulsanteID) legge la posizione
Property Get PulsantePos() As Long
PulsantePos = Pulsante(m_PulsanteID).Left
End Property
'PROP: L/S; (dopo aver assegnato un indirizzo a pulsanteID) assegna l'abilitazione
Property Let PulsanteAbilitato(ByVal Stato As Byte)
If Stato = 1 Then
   Pulsante(m_PulsanteID).Enabled = True
Else
   Pulsante(m_PulsanteID).Enabled = False
End If
PropertyChanged "PulsanteAbilitato"
UserControl.Refresh
End Property
'PROP: L/S; (dopo aver assegnato un indirizzo a pulsanteID) legge stato abilitazione
Property Get PulsanteAbilitato() As Byte
If Pulsante(m_PulsanteID).Enabled = True Then
   PulsanteAbilitato = 1
Else
   PulsanteAbilitato = 0
End If
End Property
Property Let Allarme(ByVal OnOff As Boolean)
m_Allarme = OnOff
If m_Allarme = True Then
  Timer1.Enabled = True
Else
  Timer1.Enabled = False
    If m_Selezionato <> pAllarmi Then
          Pulsante(7) = LoadResPicture(107, 0)
    Else
          Pulsante(7) = LoadResPicture(207, 0)
    End If
End If
End Property
Property Get Allarme() As Boolean
   Allarme = m_Allarme
End Property

'PROP: L/S; (dopo aver assegnato un indirizzo a pulsanteID) assegna la visibilità
Property Let PulsanteVisibile(ByVal Stato As Byte)
If Stato = 1 Then
   Pulsante(m_PulsanteID).Visible = True
Else
   Pulsante(m_PulsanteID).Visible = False
End If
PropertyChanged "PulsanteVisibile"
UserControl.Refresh
End Property
'PROP: L/S; (dopo aver assegnato un indirizzo a pulsanteID) legge visibilità
Property Get PulsanteVisibile() As Byte
If Pulsante(m_PulsanteID).Visible = True Then
   PulsanteVisibile = 1
Else
   PulsanteVisibile = 0
End If
End Property
'PROP: L/S; assegna l'id al pulsante
Property Let PulsanteID(ByVal Index As IndicePuls)
If Index > 0 And Index < 13 Then
   m_PulsanteID = Index
Else
   Err.Raise vbObjectError + 513, "Barra opzioni", "BarraOpzioni pulsanteID:Range valore (1-12)"
End If

End Property
'PROP: L/S; legge l'id del pulsante
Property Get PulsanteID() As IndicePuls
PulsanteID = m_PulsanteID
End Property

'PROPERTY BAG: LEGGE PARAMETRI
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_PulsanteID = pOrdini
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulOrdiniPos", 0)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulOrdiniAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulOrdiniVis", 0)
m_PulsanteID = pstorico
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulStoricoPos", 1000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulStoricoAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulStoricoVis", 0)
m_PulsanteID = pMappa
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulMappaPos", 2000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulMappaAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulMappaVis", 0)
m_PulsanteID = pPacco
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulPaccoPos", 3000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulPaccoAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulPaccoVis", 0)
m_PulsanteID = pRegge
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulReggePos", 4000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulReggeAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulReggeVis", 0)
m_PulsanteID = pPesa
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulPesaPos", 5000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulPesaAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulPesaVis", 0)
m_PulsanteID = pAllarmi
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulAllarmiPos", 6000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulAllarmiAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulAllarmiVis", 0)
m_PulsanteID = pEntrata
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulEntrataPos", 7000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulEntrataAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulEntrataVis", 0)
m_PulsanteID = pFilettatura
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulFilettaturaPos", 8000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulFilettaturaAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulFilettaturaVis", 0)
m_PulsanteID = pSmussatrice
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulSmussatricePos", 9000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulSmussatriceAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulSmussatriceVis", 0)
m_PulsanteID = pTaglio
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulTaglioPos", 10000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulTaglioAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulTaglioVis", 0)
m_PulsanteID = pWalkingBeam
Pulsante(m_PulsanteID).Left = PropBag.ReadProperty("PulWalkingBeamPos", 11000)
Pulsante(m_PulsanteID).Enabled = PropBag.ReadProperty("PulWalkingBeamAbi", 0)
Pulsante(m_PulsanteID).Visible = PropBag.ReadProperty("PulWalkingBeamVis", 0)
End Sub
'PROPERTY BAG: scrittura parametri
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
m_PulsanteID = pOrdini
Call PropBag.WriteProperty("PulOrdiniPos", Pulsante(m_PulsanteID).Left, 0)
Call PropBag.WriteProperty("PulOrdiniAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulOrdiniVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pstorico
Call PropBag.WriteProperty("PulStoricoPos", Pulsante(m_PulsanteID).Left, 1000)
Call PropBag.WriteProperty("PulstoricoAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulStoricoVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pMappa
Call PropBag.WriteProperty("PulMappaPos", Pulsante(m_PulsanteID).Left, 2000)
Call PropBag.WriteProperty("PulMappaAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulMappaVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pPacco
Call PropBag.WriteProperty("PulPaccoPos", Pulsante(m_PulsanteID).Left, 3000)
Call PropBag.WriteProperty("PulPaccoAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulPaccoVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pRegge
Call PropBag.WriteProperty("PulreggePos", Pulsante(m_PulsanteID).Left, 4000)
Call PropBag.WriteProperty("PulReggeAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulReggeVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pPesa
Call PropBag.WriteProperty("PulPesaPos", Pulsante(m_PulsanteID).Left, 5000)
Call PropBag.WriteProperty("PulPesaAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulPesaVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pAllarmi
Call PropBag.WriteProperty("PulAllarmiPos", Pulsante(m_PulsanteID).Left, 6000)
Call PropBag.WriteProperty("PulAllarmiAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulAllarmiVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pEntrata
Call PropBag.WriteProperty("PulEntrataPos", Pulsante(m_PulsanteID).Left, 7000)
Call PropBag.WriteProperty("PulEntrataAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulEntrataVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pFilettatura
Call PropBag.WriteProperty("PulFilettaturaPos", Pulsante(m_PulsanteID).Left, 8000)
Call PropBag.WriteProperty("PulFilettaturaAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulFilettaturaVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pSmussatrice
Call PropBag.WriteProperty("PulSmussatricePos", Pulsante(m_PulsanteID).Left, 9000)
Call PropBag.WriteProperty("PulSmussatriceAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulSmussatriceVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pTaglio
Call PropBag.WriteProperty("PulTaglioPos", Pulsante(m_PulsanteID).Left, 10000)
Call PropBag.WriteProperty("PulTaglioAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulTaglioVis", Pulsante(m_PulsanteID).Visible, 0)
m_PulsanteID = pWalkingBeam
Call PropBag.WriteProperty("PulWalkingBeamPos", Pulsante(m_PulsanteID).Left, 11000)
Call PropBag.WriteProperty("PulWalkingBeamAbi", Pulsante(m_PulsanteID).Enabled, 0)
Call PropBag.WriteProperty("PulWalkingBeamVis", Pulsante(m_PulsanteID).Visible, 0)
End Sub
