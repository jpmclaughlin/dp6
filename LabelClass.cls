VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabelClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum TCampoLingua
     Ita = 0 ' "ITALIANO"
     Ing = 1 ' "INGLESE"
     Fra = 2 ' "FRANCESE"
     Ted = 3 ' "TEDESCO"
     Spa = 4 ' "SPAGNOLO"
     Spl = 5 ' "LinguaSpeciale"
End Enum

Enum TUscita
        Video = 1
        Stampante = 2
End Enum

Enum TUnitMis
        metrica = 0
        inch = 1
End Enum

Private Type Dati
        dIntestazioneNome(3) As String
        dIntestazioneBold(3) As Boolean
        dIntestazioneFont(3) As String
        dIntestazioneSize(3) As Integer
        dIntestazionePosX(3) As Integer
        dIntestazionePosY(3) As Integer
        dCampoFisso(20) As String
        dCampoManuale(10) As String
        dCampoAuto(10) As String
        dNomeFontFisso(20) As String
        dNomeFontCampi(20) As String
        dBoldFissi(20) As Boolean
        dBoldCampi(20) As Boolean
        dSizeFissi(20) As Integer
        dSizeCampi(20) As Integer
        dPosXFissi(20) As Integer
        dPosXCampi(20) As Integer
        dPosYFissi(20) As Integer
        dPosYCampi(20) As Integer
        dLingua As Integer
        dUnitMisura As TUnitMis
        dNomeFileSfondo As String
        dLarghezzaSfondo As Integer
        dAltezzaSfondo As Integer
        dFissoVisibile(20) As Boolean
End Type

' campi ad aggiornamento automatico del cartellino
Public Descrizione As String ' descrizione ordine
Public NumeroPacco As String
Public NumeroTubi As String
Public Data As String
Public Ora As String
Public DimensioniTubo As String
Public LunghezzaTubo As String
Public SpessoreTubo As String
Public PesoPacco As String
' campi ad aggiornamento manuale del cartellino
Private PrivPrinterData(10) As String
Private privIndice(20) As Integer
Private DatiLabel As Dati

Property Get ColonnaLingua(ByVal inLingua As TCampoLingua) As String
   Select Case inLingua
   Case 0
       ColonnaLingua = "ITALIANO"
   Case 1
       ColonnaLingua = "INGLESE"
   Case 2
       ColonnaLingua = "FRANCESE"
   Case 3
       ColonnaLingua = "TEDESCO"
   Case 4
       ColonnaLingua = "SPAGNOLO"
   Case 5
       ColonnaLingua = "LinguaSpeciale"
   End Select
End Property
Property Let manIndice(Index As Integer, value As Integer)
  If Index >= 0 And Index <= 20 Then
     privIndice(Index) = value
  End If
End Property

Property Get manIndice(Index As Integer) As Integer
   manIndice(Index) = privIndice(Index)
End Property

Property Let UnitMisura(value As TUnitMis)
   DatiLabel.dUnitMisura = value
End Property

Property Get UnitMisura() As TUnitMis
   UnitMisura = DatiLabel.dUnitMisura
End Property

Property Let IntestazioneBold(Index As Integer, value As Boolean)
  If Index >= 1 And Index <= 3 Then
   DatiLabel.dIntestazioneBold(Index) = value
  End If
End Property

Property Get IntestazioneBold(Index As Integer) As Boolean
   IntestazioneBold(Index) = DatiLabel.dIntestazioneBold(Index)
End Property
Property Let IntestazioneSize(Index As Integer, value As Integer)
 If Index >= 1 And Index <= 3 Then
   DatiLabel.dIntestazioneSize(Index) = value
 End If
End Property

Property Get IntestazioneSize(Index As Integer) As Integer
   IntestazioneSize(Index) = DatiLabel.dIntestazioneSize(Index)
End Property
Property Let IntestazionePosY(Index As Integer, value As Integer)
  If Index >= 1 And Index <= 3 Then
   DatiLabel.dIntestazionePosY(Index) = value
  End If
End Property

Property Get IntestazionePosY(Index As Integer) As Integer
   IntestazionePosY(Index) = DatiLabel.dIntestazionePosY(Index)
End Property
Property Let IntestazionePosX(Index As Integer, value As Integer)
  If Index >= 1 And Index <= 3 Then
   DatiLabel.dIntestazionePosX(Index) = value
  End If
End Property

Property Get IntestazionePosX(Index As Integer) As Integer
   IntestazionePosX(Index) = DatiLabel.dIntestazionePosX(Index)
End Property

Property Let IntestazioneNome(Index As Integer, value As String)
  If Index >= 1 And Index <= 3 Then
   DatiLabel.dIntestazioneNome(Index) = value
  End If
End Property

Property Get IntestazioneNome(Index As Integer) As String
   IntestazioneNome(Index) = DatiLabel.dIntestazioneNome(Index)
End Property
Property Let IntestazioneFont(Index As Integer, value As String)
   If Index >= 1 And Index <= 3 Then
        DatiLabel.dIntestazioneFont(Index) = value
   End If
End Property

Property Get IntestazioneFont(Index As Integer) As String
   IntestazioneFont(Index) = DatiLabel.dIntestazioneFont(Index)
End Property
Property Let SfondoFilenome(value As String)
        DatiLabel.dNomeFileSfondo = value
   End Property

Property Get SfondoFilenome() As String
   SfondoFilenome = DatiLabel.dNomeFileSfondo
End Property
Property Let SfondoAltezza(value As Integer)
        DatiLabel.dAltezzaSfondo = value
   End Property

Property Get SfondoAltezza() As Integer
   SfondoAltezza = DatiLabel.dAltezzaSfondo
End Property
Property Let SfondoLarghezza(value As Integer)
        DatiLabel.dLarghezzaSfondo = value
   End Property

Property Get SfondoLarghezza() As Integer
   SfondoLarghezza = DatiLabel.dLarghezzaSfondo
End Property

Property Let FileSfondo(value As String)
        DatiLabel.dNomeFileSfondo = value
   End Property

Property Get FileSfondo() As String
   FileSfondo = DatiLabel.dNomeFileSfondo
End Property

Property Let Lingua(value As TCampoLingua)
     DatiLabel.dLingua = value
     CaricaCampiFissi (value)
End Property

Property Get Lingua() As TCampoLingua
     Lingua = DatiLabel.dLingua
End Property

Property Let CampiPosY(Index As Integer, value As Integer)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dPosYCampi(Index) = value
    End If
End Property

Property Get CampiPosY(Index As Integer) As Integer
    If Index >= 1 And Index <= 20 Then
        CampiPosY = DatiLabel.dPosYCampi(Index)
    Else
        CampiPosY = 0
    End If
End Property
Property Let FissiPosY(Index As Integer, value As Integer)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dPosYFissi(Index) = value
    End If
End Property

Property Get FissiPosY(Index As Integer) As Integer
    If Index >= 1 And Index <= 20 Then
        FissiPosY = DatiLabel.dPosYFissi(Index)
    Else
        FissiPosY = 0
    End If
End Property
Property Let FissoVisibile(Index As Integer, value As Boolean)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dFissoVisibile(Index) = value
    End If
End Property

Property Get FissoVisibile(Index As Integer) As Boolean
    If Index >= 1 And Index <= 20 Then
        FissoVisibile = DatiLabel.dFissoVisibile(Index)
    Else
        FissoVisibile = False
    End If
End Property

Property Let CampiPosX(Index As Integer, value As Integer)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dPosXCampi(Index) = value
    End If
End Property

Property Get CampiPosX(Index As Integer) As Integer
    If Index >= 1 And Index <= 20 Then
        CampiPosX = DatiLabel.dPosXCampi(Index)
    Else
        CampiPosX = 0
    End If
End Property
Property Let FissiPosX(Index As Integer, value As Integer)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dPosXFissi(Index) = value
    End If
End Property

Property Get FissiPosX(Index As Integer) As Integer
    If Index >= 1 And Index <= 20 Then
        FissiPosX = DatiLabel.dPosXFissi(Index)
    Else
        FissiPosX = 0
    End If
End Property
Property Let Fisso(Index As Integer, value As String)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dCampoFisso(Index) = value
    End If
End Property

Property Get Fisso(Index As Integer) As String
    If Index >= 1 And Index <= 20 Then
        Fisso = DatiLabel.dCampoFisso(Index)
    Else
        Fisso = ""
    End If
End Property
Property Let FissiFont(Index As Integer, value As String)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dNomeFontFisso(Index) = value
    End If
End Property

Property Get FissiFont(Index As Integer) As String
    If Index >= 1 And Index <= 20 Then
        FissiFont = DatiLabel.dNomeFontFisso(Index)
    Else
        FissiFont = "Arial"
    End If
End Property
Property Let CampiFont(Index As Integer, value As String)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dNomeFontCampi(Index) = value
    End If
End Property

Property Get CampiFont(Index As Integer) As String
    If Index >= 1 And Index <= 20 Then
        CampiFont = DatiLabel.dNomeFontCampi(Index)
    Else
        CampiFont = "Arial"
    End If
End Property

Property Let CampiSize(Index As Integer, value As Integer)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dSizeCampi(Index) = value
    End If
End Property

Property Get CampiSize(Index As Integer) As Integer
    If Index >= 1 And Index <= 20 Then
        CampiSize = DatiLabel.dSizeCampi(Index)
    Else
        CampiSize = 12
    End If
End Property
Property Let FissiSize(Index As Integer, value As Integer)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dSizeFissi(Index) = value
    End If
End Property

Property Get FissiSize(Index As Integer) As Integer
    If Index >= 1 And Index <= 20 Then
        FissiSize = DatiLabel.dSizeFissi(Index)
    Else
        FissiSize = 12
    End If
End Property

Property Let FissiBold(Index As Integer, value As Boolean)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dBoldFissi(Index) = value
    End If
End Property

Property Get FissiBold(Index As Integer) As Boolean
    If Index >= 1 And Index <= 20 Then
        FissiBold = DatiLabel.dBoldFissi(Index)
    Else
        FissiBold = False
    End If
End Property

Property Let CampiBold(Index As Integer, value As Boolean)
    If Index >= 1 And Index <= 20 Then
        DatiLabel.dBoldCampi(Index) = value
    End If
End Property

Property Get CampiBold(Index As Integer) As Boolean
    If Index >= 1 And Index <= 20 Then
        CampiBold = DatiLabel.dCampoAuto(Index)
    Else
        CampiBold = False
    End If
End Property
Property Let CampoAuto(Index As Integer, value As String)
    If Index >= 1 And Index <= 10 Then
        DatiLabel.dCampoAuto(Index) = value
    End If
End Property

Property Get CampoAuto(Index As Integer) As String
    If Index >= 1 And Index <= 10 Then
        CampoAuto = DatiLabel.dCampoAuto(Index)
    Else
        CampoAuto = ""
    End If
End Property
Property Let CampoManuale(Index As Integer, value As String)
    If Index >= 1 And Index <= 10 Then
        'PrivPrinterData(Index) = Value
        DatiLabel.dCampoManuale(Index) = value
    End If
End Property

Property Get CampoManuale(Index As Integer) As String
    If Index >= 1 And Index <= 10 Then
    '    CampoManuale = PrivPrinterData(Index)
        CampoManuale = DatiLabel.dCampoManuale(Index)
    Else
        CampoManuale = ""
    End If
End Property

Sub DrawLabel(Uscita As TUscita, Optional Dispositivo As Object)
    Dim i As Integer
    Dim Anteprima As PictureBox
    Dim a As String
    
    On Error Resume Next
    If Uscita = Video Then
    Set Anteprima = Dispositivo
    With Anteprima
        .Cls
        .Width = DatiLabel.dLarghezzaSfondo
        .Height = DatiLabel.dAltezzaSfondo
        a = BitmapPath + DatiLabel.dNomeFileSfondo + ".gif"
        .Picture = LoadPicture(a)
        For i = 1 To 20
               If DatiLabel.dFissoVisibile(i) Then
                    .FontName = DatiLabel.dNomeFontFisso(i)
                    .FontSize = DatiLabel.dSizeFissi(i)
                    .FontBold = DatiLabel.dBoldFissi(i)
                    .CurrentX = DatiLabel.dPosXFissi(i)
                    .CurrentY = DatiLabel.dPosYFissi(i)
                    Anteprima.Print DatiLabel.dCampoFisso(i)
   
                    .FontName = DatiLabel.dNomeFontCampi(i)
                    .FontSize = DatiLabel.dSizeCampi(i)
                    .FontBold = DatiLabel.dBoldCampi(i)
                    .CurrentX = DatiLabel.dPosXCampi(i)
                    .CurrentY = DatiLabel.dPosYCampi(i)
                    If i < 11 Then
                       Anteprima.Print DatiLabel.dCampoManuale(i)
                    Else
                       Anteprima.Print DatiLabel.dCampoAuto(i - 10)
                    End If
              End If
        Next

        For i = 0 To 2
                    .Font.Name = DatiLabel.dIntestazioneFont(i)
                    .Font.Size = IIf(DatiLabel.dIntestazioneSize(i) > 0, DatiLabel.dIntestazioneSize(i), 10)
                    .Font.Bold = DatiLabel.dIntestazioneBold(i)
                    .CurrentX = DatiLabel.dIntestazionePosX(i)
                    .CurrentY = DatiLabel.dIntestazionePosY(i)
                    Anteprima.Print DatiLabel.dIntestazioneNome(i)
        Next
    End With
    Set Dispositivo = Anteprima
    Else
     '   Set Anteprima = Dispositivo
      With Dispositivo
        '.Width = DatiLabel.dLarghezzaSfondo
        '.Height = DatiLabel.dAltezzaSfondo
        a = BitmapPath + DatiLabel.dNomeFileSfondo + ".gif"
        .PaintPicture LoadPicture(a), 0, 0
        For i = 1 To 20
                If DatiLabel.dFissoVisibile(i) Then
                    .FontName = DatiLabel.dNomeFontFisso(i)
                    .FontSize = DatiLabel.dSizeFissi(i)
                    .FontBold = DatiLabel.dBoldFissi(i)
                    .CurrentX = DatiLabel.dPosXFissi(i)
                    .CurrentY = DatiLabel.dPosYFissi(i)
                    If DatiLabel.dCampoFisso(i) <> "" Then Dispositivo.Print DatiLabel.dCampoFisso(i)

                    .FontName = DatiLabel.dNomeFontCampi(i)
                    .FontSize = DatiLabel.dSizeCampi(i)
                    .FontBold = DatiLabel.dBoldCampi(i)
                    .CurrentX = DatiLabel.dPosXCampi(i)
                    .CurrentY = DatiLabel.dPosYCampi(i)
                    If DatiLabel.dCampoFisso(i) <> "" Then
                    If i < 11 Then
                       Dispositivo.Print DatiLabel.dCampoManuale(i)
                    Else
                       Dispositivo.Print DatiLabel.dCampoAuto(i - 10)
                    End If
                    End If
               End If
        Next

        For i = 0 To 2
                    If DatiLabel.dIntestazioneFont(i) = "" Then
                       DatiLabel.dIntestazioneFont(i) = "Arial"
                    Else
                       .Font.Name = DatiLabel.dIntestazioneFont(i)
                    End If
                    .Font.Size = IIf(DatiLabel.dIntestazioneSize(i) > 0, DatiLabel.dIntestazioneSize(i), 10)
                    .Font.Bold = DatiLabel.dIntestazioneBold(i)
                    .CurrentX = DatiLabel.dIntestazionePosX(i)
                    .CurrentY = DatiLabel.dIntestazionePosY(i)
                    If DatiLabel.dCampoFisso(i) <> "" Then Dispositivo.Print DatiLabel.dIntestazioneNome(i)
        Next
    End With
    'Dispositivo.EndDoc
    'Set Dispositivo = Anteprima
    End If
End Sub
'----------------------------------------
' LEGGE I DATI DEL FILE PARAMETRI
'----------------------------------------
Public Function Read_file_dati() As Boolean
    Dim i As Integer
    Dim FileID
    
    On Error GoTo Errore
    Read_file_dati = False
    
    FileID = FreeFile()
    Open TargetPath & "LabelData.dat" For Input As FileID
    With DatiLabel
       Input #1, .dAltezzaSfondo
       Input #1, .dLarghezzaSfondo
       Input #1, .dLingua
       Input #1, .dNomeFileSfondo
       Input #1, .dUnitMisura
       For i = 0 To 20
          '-----------------------
          Input #1, .dBoldFissi(i)
          Input #1, .dCampoFisso(i)
          Input #1, .dNomeFontFisso(i)
          Input #1, .dPosXFissi(i)
          Input #1, .dPosYFissi(i)
          Input #1, .dSizeFissi(i)
          '-----------------------
          Input #1, .dBoldCampi(i)
          Input #1, .dNomeFontCampi(i)
          Input #1, .dPosXCampi(i)
          Input #1, .dPosYCampi(i)
          Input #1, .dSizeCampi(i)
          Input #1, .dFissoVisibile(i)
       Next
       For i = 0 To 10
          Input #1, .dCampoAuto(i)
          Input #1, .dCampoManuale(i)
       Next
       For i = 0 To 3
          Input #1, .dIntestazioneBold(i)
          Input #1, .dIntestazioneFont(i)
          Input #1, .dIntestazioneNome(i)
          Input #1, .dIntestazionePosX(i)
          Input #1, .dIntestazionePosY(i)
          Input #1, .dIntestazioneSize(i)
       Next
    End With
    Read_file_dati = True
    
Errore:
    Close FileID
  
End Function

'----------------------------------------
' SCRIVE I DATI NEL FILE PARAMETRI
'----------------------------------------
Public Function Scrive_file_dati() As Boolean
    Dim i As Integer
    Dim FileID
     
    On Error GoTo Errore
    Scrive_file_dati = False
    FileID = FreeFile()
    Open TargetPath & "LabelData.dat" For Output As FileID
    With DatiLabel
       Write #1, .dAltezzaSfondo
       Write #1, .dLarghezzaSfondo
       Write #1, .dLingua
       Write #1, .dNomeFileSfondo
       Write #1, .dUnitMisura
       For i = 0 To 20
          '-----------------------
          Write #1, .dBoldFissi(i)
          Write #1, .dCampoFisso(i)
          Write #1, .dNomeFontFisso(i)
          Write #1, .dPosXFissi(i)
          Write #1, .dPosYFissi(i)
          Write #1, .dSizeFissi(i)
          '-----------------------
          Write #1, .dBoldCampi(i)
          Write #1, .dNomeFontCampi(i)
          Write #1, .dPosXCampi(i)
          Write #1, .dPosYCampi(i)
          Write #1, .dSizeCampi(i)
          Write #1, .dFissoVisibile(i)
       Next
       For i = 0 To 10
          Write #1, .dCampoAuto(i)
          Write #1, .dCampoManuale(i)
       Next
       For i = 0 To 3
          Write #1, .dIntestazioneBold(i)
          Write #1, .dIntestazioneFont(i)
          Write #1, .dIntestazioneNome(i)
          Write #1, .dIntestazionePosX(i)
          Write #1, .dIntestazionePosY(i)
          Write #1, .dIntestazioneSize(i)
       Next
    End With
    Scrive_file_dati = True
    
Errore:
    Close FileID
End Function

Sub CaricaCampiFissi(CampoLingua As TCampoLingua)
    Dim i As Byte
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim StrL As String
    
    Select Case CampoLingua
    Case 1
         StrL = "ITALIANO"
    Case 2
         StrL = "INGLESE"
    Case 3
         StrL = "FRANCESE"
     Case 4
         StrL = "TEDESCO"
     Case 5
         StrL = "SPAGNOLO"
     Case 6
         StrL = "LinguaSpeciale"
    End Select
    
    On Error Resume Next
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\target\production.mdb;Persist Security Info=False"
    With rs
        .CursorLocation = adUseClient
        .Open "SELECT * FROM TestiCartellino", cn, adOpenStatic, adLockReadOnly, adCmdText
        Set .ActiveConnection = Nothing
        cn.Close
        ' chiude la connessione, rimane aperto il recordset
        If .EOF = False Then
            .Find ("TagName='Intestaz01'")
            DatiLabel.dIntestazioneNome(1) = .Fields(StrL)
            .Find ("TagName='Intestaz02'")
            DatiLabel.dIntestazioneNome(2) = .Fields(StrL)
            .Find ("TagName='Intestaz03'")
            DatiLabel.dIntestazioneNome(3) = .Fields(StrL)
            For i = 1 To 20
               .MoveFirst
               .Find ("TagName='Campo0" & LTrim$(i) & "'")
               DatiLabel.dCampoFisso(i) = .Fields(StrL)
            Next
        End If
        ' chiude il recordset
        .Close
    End With
    Set rs = Nothing
    Set cn = Nothing
End Sub
