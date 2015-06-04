VERSION 5.00
Begin VB.UserControl Allarme 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   ScaleHeight     =   1635
   ScaleWidth      =   1650
   ToolboxBitmap   =   "Allarme.ctx":0000
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ALLARME"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuDefinizione 
         Caption         =   "Definizione"
      End
   End
End
Attribute VB_Name = "Allarme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' enumerazione tipi predefiniti

Enum AllTipo
     Nessuno = 0
     Allarme = 1
     Messaggio = 2
     Entrambi = 3
End Enum

' dichiarazione variabili

Private m_Tipo As AllTipo
Private m_Intervallo As Integer

'dichiarazione evento

Public Event Cliccato()
'genera evento "cliccato" con image_click
Private Sub Image1_Click()
RaiseEvent Cliccato
End Sub

Sub RefreshTimer()
' controllo stato variabili
Static anima As Boolean

If IsEmpty(anima) = True Then anima = False
If m_Tipo = 0 Or IsEmpty(m_Tipo) = True Then m_Tipo = Allarme
If IsEmpty(m_Tipo) = True Or IsNull(m_Tipo) = True Then m_Tipo = Nessuno

' controllo stato animazione

If anima = False Then
   Image1.Picture = LoadResPicture(404, 0)
Else
   Image1.Picture = LoadResPicture(400 + m_Tipo, 0)
End If

' cambio stato animazione

anima = Not (anima)
'
If m_Tipo <> Nessuno Then
'   Timer1.Enabled = True
   Image1.Visible = True
Else
'   Timer1.Enabled = False
   Image1.Visible = False
End If
End Sub

' INIZIALIZZAZIONE controllo

Private Sub UserControl_Initialize()
If (m_Tipo <> Nessuno) Then
  ' Timer1.Enabled = True
   Image1.Visible = True
Else
   'Timer1.Enabled = False
   Image1.Visible = False
End If

'Timer1.Interval = 1000
'anima = False
End Sub

'LET: assegna il tipo di allarme

Property Let AllarmeTipo(Tp As AllTipo)
m_Tipo = Tp

If m_Tipo <> Nessuno Then
  ' Timer1.Enabled = True
   Image1.Visible = True
Else
  ' Timer1.Enabled = False
   Image1.Visible = False
End If
'anima = False
End Property

'GET: legge il tipo di allarme

Property Get AllarmeTipo() As AllTipo
  AllarmeTipo = m_Tipo
End Property

'LET: imposta l'intervallo del timer di lampeggio

Property Let Intervallo(Ingresso As Integer)
If Ingresso > 1 And Ingresso < 9000 Then
   'Timer1.Interval = Ingresso
   m_Intervallo = Ingresso
Else
   ' generazione errore se il valore del timer supera i limiti impostati
   
   Err.Raise vbObjectError + 513, "Segnalazioni.Timer1", _
             "Intervallo " & Ingresso & " non valido"
End If
End Property

'GET: legge intervallo del timer

Property Get Intervallo() As Integer
Intervallo = m_Intervallo
End Property

Private Sub UserControl_Resize()

'controllo se l'ambiente è in modalità progettazione o esecuzione

If Ambient.UserMode = True Then
   Shape1.Visible = False
   Label1.Visible = False
 '  Timer1.Enabled = True
Else
   'Timer1.Enabled = False
   Shape1.Visible = True
   Label1.Visible = True
End If
End Sub
