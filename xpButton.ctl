VERSION 5.00
Begin VB.UserControl XPButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   DefaultCancel   =   -1  'True
   FillColor       =   &H00B1B1B1&
   ScaleHeight     =   48
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   190
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   2745
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   150
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNDKSHADOW = 21
Private Const COLOR_BTNLIGHT = 22

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTERABS = &H65

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF = 4

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum ButtonTypes
    [Windows 16-bit] = 1
    [Windows 32-bit] = 2
    [Windows XP chrome] = 3
    [Mac] = 4
    [Java metal] = 5
    [Netscape 6] = 6
    [Simple Flat] = 7
End Enum

Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
    [Force Standard] = 3
End Enum

'eventi
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

'variabili
Private MyButtonType As ButtonTypes
Private MyColorType As ColorTypes

Private He As Long          'height button
Private Wi As Long          'width  button
Private BackC As Long       'backColor
Private ForeC As Long       'foreColor
Private elTex As String     'current text
Private TextFont As StdFont 'current font

Private rc As RECT, rc2 As RECT, rc3 As RECT
Private rgnNorm As Long
Private LastButton As Byte
Private isEnabled As Boolean
Private hasFocus As Boolean, showFocusR As Boolean
Private cFace As Long, cLight As Long, cHighLight As Long, cShadow As Long, cDarkShadow As Long, cText As Long
Private lastStat As Byte, TE As String
'=======================
Private m_TxtTop As Integer
Private m_TxtLeft As Integer
Private m_TxtText As String
'=======================
Private m_ImgTop As Integer
Private m_ImgLeft As Integer
Private ImgPath As String
Private m_ImgAllarga As Boolean
Private m_imgWidth As Integer
Private m_imgHeight As Integer

Property Let TxtText(ByVal NewValue As String)
   m_TxtText = NewValue
   Label1.Caption = m_TxtText
End Property
Property Get TxtText() As String
   TxtText = m_TxtText
End Property

Property Let TxtTop(ByVal NewValue As Integer)
   m_TxtTop = NewValue
   Label1.Top = m_TxtTop
End Property
Property Get TxtTop() As Integer
   TxtTop = m_TxtTop
End Property
Property Let TxtLeft(ByVal NewValue As Integer)
   m_TxtLeft = NewValue
   Label1.Left = m_TxtLeft
End Property
Property Get TxtLeft() As Integer
   TxtLeft = m_TxtLeft
End Property
Property Let ImgH(ByVal NewValue As Integer)
   m_imgHeight = NewValue
   Image2.Height = m_imgHeight
End Property
Property Get ImgH() As Integer
   ImgH = m_imgHeight
End Property
Property Let ImgW(ByVal NewValue As Integer)
   m_imgWidth = NewValue
   Image2.Width = m_imgWidth
End Property
Property Get ImgW() As Integer
   ImgW = m_imgWidth
End Property

Property Let ImgAllarga(ByVal NewValue As Boolean)
   m_ImgAllarga = NewValue
   Image2.Stretch = m_ImgAllarga
End Property
Property Get ImgAllarga() As Boolean
   ImgAllarga = m_ImgAllarga
End Property

Property Let ImgTop(ByVal NewValue As Integer)
   m_ImgTop = NewValue
   Image2.Top = m_ImgTop
End Property
Property Get ImgTop() As Integer
   ImgTop = m_ImgTop
End Property
Property Let ImgLeft(ByVal NewValue As Integer)
   m_ImgLeft = NewValue
   Image2.Left = m_ImgLeft
End Property
Property Get ImgLeft() As Integer
   ImgLeft = m_ImgLeft
End Property
Property Let Icona(ByVal Percorso As String)
   ImgPath = Percorso
   Image2.Picture = LoadPicture(ImgPath)
End Property
Property Get Icona() As String
   Icona = ImgPath
End Property

Private Sub Image2_Click()
    Call UserControl_Click
    
'    If (LastButton = 1) And (isEnabled = True) Then
'     Call Redraw(0, True) 'normal status drawn
'     UserControl.Refresh
'     RaiseEvent Click
'  '   Call UserControl_Click
'  End If
End Sub

Private Sub Label1_Click()
   Call UserControl_Click
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  Call UserControl_Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Click()
  If (LastButton = 1) And (isEnabled = True) Then
     Call Redraw(0, True) 'normal status drawn
     UserControl.Refresh
     RaiseEvent Click
  End If
End Sub

Private Sub UserControl_DblClick()
  If LastButton = 1 Then
     Call UserControl_MouseDown(1, 1, 1, 1)
  End If
End Sub

Private Sub UserControl_GotFocus()
  hasFocus = True
  Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 32 Then 'barra spazio premuta
     Call UserControl_MouseDown(1, 1, 1, 1)
  ElseIf (KeyCode = 39) Or (KeyCode = 40) Then 'rightdown
     SendKeys "{Tab}"
  ElseIf (KeyCode = 37) Or (KeyCode = 38) Then 'leftup
     SendKeys "+{Tab}"
  End If
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 32 Then 'spacebar
     Call UserControl_MouseUp(1, 1, 1, 1)
     LastButton = 1
     Call UserControl_Click
  End If
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
  hasFocus = False
  Call Redraw(lastStat, True)
End Sub

Private Sub UserControl_Initialize()
  LastButton = 1
  rc2.Left = 2: rc2.Top = 2
  Call SetColors
End Sub

Private Sub UserControl_InitProperties()
  isEnabled = True
  showFocusR = True
  Set TextFont = UserControl.Font
  MyButtonType = [Windows 32-bit]
  ImgTop = 5
  ImgLeft = 5
  ImgH = 10
  ImgW = 10
  m_TxtLeft = 5
  m_TxtTop = 5
  ImgAllarga = False
  TxtText = ""
  ImgPath = ""
  Icona = ""
  MyColorType = [Use Windows]
  BackC = GetSysColor(COLOR_BTNFACE)
  ForeC = GetSysColor(COLOR_BTNTEXT)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  LastButton = Button
  If Button <> 2 Then Call Redraw(2, False)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button < 2 Then
     If x < 0 Or y < 0 Or x > Wi Or y > He Then
        'outside button
        Call Redraw(0, False)
     Else
        'inside button
        If Button = 1 Then Call Redraw(2, False)
     End If
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button <> 2 Then Call Redraw(0, False)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'*******************************************
'PROPRITA CONTROLLO                        *
'*******************************************
Public Property Get BackColor() As OLE_COLOR
  BackColor = BackC
End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)
  BackC = theCol
  Call SetColors
  Call Redraw(lastStat, True)
  PropertyChanged "BCOL"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = ForeC
End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)
  ForeC = theCol
  Call SetColors
  Call Redraw(lastStat, True)
  PropertyChanged "FCOL"
End Property

Public Property Get ButtonType() As ButtonTypes
  ButtonType = MyButtonType
End Property

Public Property Let ButtonType(ByVal NewValue As ButtonTypes)
  MyButtonType = NewValue
  Call Redraw(0, True)
  PropertyChanged "BTYPE"
End Property

Public Property Get Caption() As String
  Caption = elTex
End Property

Public Property Let Caption(ByVal NewValue As String)
  elTex = NewValue
  Call SetAccessKeys
  Call Redraw(0, True)
  PropertyChanged "TX"
End Property

Public Property Get Enabled() As Boolean
  Enabled = isEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  isEnabled = NewValue
  Call Redraw(0, True)
  UserControl.Enabled = isEnabled
  PropertyChanged "ENAB"
End Property

Public Property Get Font() As Font
  Set Font = TextFont
End Property

Public Property Set Font(ByRef newFont As Font)
  Set TextFont = newFont
  Set UserControl.Font = TextFont
  Call Redraw(0, True)
  PropertyChanged "FONT"
End Property

Public Property Get ColorScheme() As ColorTypes
  ColorScheme = MyColorType
End Property

Public Property Let ColorScheme(ByVal NewValue As ColorTypes)
  MyColorType = NewValue
  Call SetColors
  Call Redraw(0, True)
  PropertyChanged "COLTYPE"
End Property

Public Property Get ShowFocusRect() As Boolean
  ShowFocusRect = showFocusR
End Property

Public Property Let ShowFocusRect(ByVal NewValue As Boolean)
  showFocusR = NewValue
  Call Redraw(lastStat, True)
  PropertyChanged "FOCUSR"
End Property

Public Property Get hwnd() As Long
  hwnd = UserControl.hwnd
End Property

'**********************************************
'FINE PROPRIETA                               *
'**********************************************
Private Sub UserControl_Resize()
  He = UserControl.ScaleHeight
  Wi = UserControl.ScaleWidth
  rc.Bottom = He: rc.Right = Wi
  rc2.Bottom = He: rc2.Right = Wi
  rc3.Left = 4: rc3.Top = 4: rc3.Right = Wi - 4: rc3.Bottom = He - 4
  DeleteObject rgnNorm
  Call MakeRegion
  SetWindowRgn UserControl.hwnd, rgnNorm, True
  Call Redraw(0, True)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  
  TxtText = PropBag.ReadProperty("TxtText", " ")
  TxtTop = PropBag.ReadProperty("TxtTop", 5)
  TxtLeft = PropBag.ReadProperty("TxtLeft", 5)
  ImgTop = PropBag.ReadProperty("IMGTOP", 5)
  ImgLeft = PropBag.ReadProperty("IMGLEFT", 5)
  Icona = PropBag.ReadProperty("ICONA", "")
  ImgH = PropBag.ReadProperty("ImgH", 10)
  ImgW = PropBag.ReadProperty("ImgW", 10)
  ImgAllarga = PropBag.ReadProperty("ImgAllarga", False)
  MyButtonType = PropBag.ReadProperty("BTYPE", 2)
  elTex = PropBag.ReadProperty("TX", "")
  isEnabled = PropBag.ReadProperty("ENAB", True)
  Set TextFont = PropBag.ReadProperty("FONT", UserControl.Font)
  MyColorType = PropBag.ReadProperty("COLTYPE", 1)
  showFocusR = PropBag.ReadProperty("FOCUSR", True)
  BackC = PropBag.ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
  ForeC = PropBag.ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
  UserControl.Enabled = isEnabled
  Set UserControl.Font = TextFont
  Call SetColors
  Call SetAccessKeys
  Call Redraw(0, True)
End Sub



Private Sub UserControl_Terminate()
  DeleteObject rgnNorm
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("TxtText", TxtText)
  Call PropBag.WriteProperty("TxtTop", TxtTop)
  Call PropBag.WriteProperty("TxtLeft", TxtLeft)
  Call PropBag.WriteProperty("BTYPE", MyButtonType)
  Call PropBag.WriteProperty("IMGTOP", ImgTop)
  Call PropBag.WriteProperty("IMGLEFT", ImgLeft)
  Call PropBag.WriteProperty("ICONA", Icona)
  Call PropBag.WriteProperty("ImgW", ImgW)
  Call PropBag.WriteProperty("ImgH", ImgH)
  Call PropBag.WriteProperty("ImgAllarga", ImgAllarga)
  Call PropBag.WriteProperty("TX", elTex)
  Call PropBag.WriteProperty("ENAB", isEnabled)
  Call PropBag.WriteProperty("FONT", TextFont)
  Call PropBag.WriteProperty("COLTYPE", MyColorType)
  Call PropBag.WriteProperty("FOCUSR", showFocusR)
  Call PropBag.WriteProperty("BCOL", BackC)
  Call PropBag.WriteProperty("FCOL", ForeC)
End Sub

Private Sub Redraw(ByVal curStat As Byte, ByVal Force As Boolean)
  
' Disegna il controllo in base allo stile selezionato
  If Force = False Then
     If (curStat = lastStat) And (TE = elTex) Then Exit Sub
  End If
  If He = 0 Then Exit Sub 'we don't want errors
  lastStat = curStat
  TE = elTex
  Dim i As Long, stepXP1 As Single, stepXP2 As Single
  Dim preFocusValue As Boolean
  preFocusValue = hasFocus 'save this value to restore it later
  If hasFocus = True Then hasFocus = ShowFocusRect
  With UserControl
     .Cls
     DrawRectangle 0, 0, Wi, He, cFace
     If isEnabled = True Then
        SetTextColor .hdc, cText 'restore font color
        If curStat = 0 Then
          'BUTTON NORMAL STATE
          Select Case MyButtonType
            Case 1 'Windows 16-bit
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
                UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
                UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
                UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                DrawRectangle 2, 2, Wi - 4, He - 4, cHighLight, True
                UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cShadow
                UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
                UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cShadow
                If hasFocus = True Then DrawFocusRect .hdc, rc3
            Case 2 'Windows 32-bit
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                If Ambient.DisplayAsDefault = True Then
                    DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
                    DrawRectangle 2, 2, Wi - 4, He - 4, cLight, True
                    UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cDarkShadow
                    UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cShadow
                    UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cDarkShadow
                    UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cShadow
                    If hasFocus = True Then DrawFocusRect .hdc, rc3
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                Else
                    DrawRectangle 0, 0, Wi - 1, He - 1, cHighLight, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, cLight, True
                    UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cDarkShadow
                    UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                    UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cDarkShadow
                    UserControl.Line (1, He - 2)-(Wi - 2, He - 2), cShadow
                End If
            Case 3 'Windows XP
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                stepXP1 = 75 / He
                stepXP2 = 50 / He
                For i = 1 To He
                    UserControl.Line (0, i)-(Wi, i), &HFFFFFF - RGB(stepXP1 * i, stepXP1 * i, stepXP2 * i)
                Next
                Label1.Top = m_TxtTop: Label1.Left = m_TxtLeft
                Label1.Caption = m_TxtText
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                UserControl.Line (2, 0)-(Wi - 2, 0), &H733C00
                UserControl.Line (2, He - 1)-(Wi - 2, He - 1), &H733C00
                UserControl.Line (0, 2)-(0, He - 2), &H733C00
                UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), &H733C00
                SetPixel .hdc, 1, 1, &H7B4D10
                SetPixel .hdc, 1, He - 2, &H7B4D10
                SetPixel .hdc, Wi - 2, 1, &H7B4D10
                SetPixel .hdc, Wi - 2, He - 2, &H7B4D10
                DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                UserControl.Line (2, He - 2)-(Wi - 2, He - 2), &HEF826B
                UserControl.Line (2, 1)-(Wi - 2, 1), &HFFE7CE
                UserControl.Line (1, 2)-(Wi - 1, 2), &HF7D7BD
                If hasFocus = True Then
                    UserControl.Line (2, 3)-(2, He - 3), &HFFFFFF
                    UserControl.Line (Wi - 3, 3)-(Wi - 3, He - 3), &HFFFFFF
                End If
            Case 4 'Mac
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
                UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
                UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
                UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
                SetPixel .hdc, 1, 1, cDarkShadow
                SetPixel .hdc, 1, He - 2, cDarkShadow
                SetPixel .hdc, Wi - 2, 1, cDarkShadow
                SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
                SetPixel .hdc, 1, 2, cFace
                SetPixel .hdc, 2, 1, cFace
                UserControl.Line (3, 2)-(Wi - 3, 2), cHighLight
                UserControl.Line (2, 2)-(2, He - 3), cHighLight
                SetPixel .hdc, 3, 3, cHighLight
                UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cFace
                UserControl.Line (1, He - 3)-(Wi - 3, He - 3), cFace
                SetPixel .hdc, Wi - 4, He - 4, cFace
                UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), cShadow
                UserControl.Line (3, He - 2)-(Wi - 2, He - 2), cShadow
                SetPixel .hdc, Wi - 3, He - 3, cShadow
                SetPixel .hdc, 2, He - 2, cFace
                SetPixel .hdc, 2, He - 3, cLight
                SetPixel .hdc, Wi - 2, 2, cFace
                SetPixel .hdc, Wi - 3, 2, cLight
            Case 5 'Java
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                .FontBold = True
                DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                SetPixel .hdc, 1, He - 2, ShiftColor(cShadow, &H1A)
                SetPixel .hdc, Wi - 2, 1, ShiftColor(cShadow, &H1A)
                If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                .FontBold = TextFont.Bold
            Case 6 'Netscape
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, ShiftColor(cLight, &H8), True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cLight, &H8), True
                UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
                UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
                UserControl.Line (0, He - 1)-(Wi, He - 1), cShadow
                UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
                If hasFocus = True Then DrawFocusRect .hdc, rc3
             Case 7 'Flat
                Image2.Top = ImgTop
                Image2.Left = ImgLeft
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cHighLight, True
                UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
                UserControl.Line (0, He - 1)-(Wi, He - 1), cShadow
                If hasFocus = True Then DrawFocusRect .hdc, rc3
        End Select
    ElseIf curStat = 2 Then
        'BUTTON DOWN #@#@#@#@#@#
        Select Case MyButtonType
            Case 1 'Windows 16-bit
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
                UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
                UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
                UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                DrawRectangle 2, 2, Wi - 4, He - 4, cShadow, True
                UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cHighLight
                UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cHighLight
                UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cHighLight
                UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cHighLight
                If hasFocus = True Then DrawFocusRect .hdc, rc3
            Case 2 'Windows 32-bit
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                If hasFocus = True Then DrawFocusRect .hdc, rc3
            Case 3 'Windows XP
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                stepXP1 = 85 / He
                stepXP2 = 60 / He
                For i = 3 To He
                    UserControl.Line (0, i)-(Wi, i), &HFFFFFF - RGB(stepXP1 * i, stepXP1 * i, stepXP2 * i)
                Next
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                UserControl.Line (2, 0)-(Wi - 2, 0), &H733C00
                UserControl.Line (2, He - 1)-(Wi - 2, He - 1), &H733C00
                UserControl.Line (0, 2)-(0, He - 2), &H733C00
                UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), &H733C00
                SetPixel .hdc, 1, 1, &H7B4D10
                SetPixel .hdc, 1, He - 2, &H7B4D10
                SetPixel .hdc, Wi - 2, 1, &H7B4D10
                SetPixel .hdc, Wi - 2, He - 2, &H7B4D10
                'DrawRectangle 1, 2, Wi - 2, He - 4, &H31B2FF, True
                DrawRectangle 1, 2, Wi - 2, He - 4, extColor.vblightBlue, True
                UserControl.Line (2, He - 2)-(Wi - 2, He - 2), extColor.vbDarkGray '&H96E7& ok
                UserControl.Line (2, 1)-(Wi - 2, 1), vbBlack  '&HCEF3FF**
                UserControl.Line (1, 2)-(Wi - 1, 2), extColor.vblightgray 'vbWhite '&H8CDBFF ok
                UserControl.Line (2, 3)-(2, He - 3), extColor.vblightgray 'vbWhite '&H6BCBFF ok
                UserControl.Line (Wi - 3, 3)-(Wi - 3, He - 3), vbWhite  '&H6BCBFF 'ok
            Case 4 'Mac
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                SetTextColor .hdc, cLight
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
                UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
                UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
                UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                SetPixel .hdc, 2, 2, ShiftColor(cShadow, -&H40)
                SetPixel .hdc, 3, 3, ShiftColor(cShadow, -&H20)
                SetPixel .hdc, 1, 1, cDarkShadow
                SetPixel .hdc, 1, He - 2, cDarkShadow
                SetPixel .hdc, Wi - 2, 1, cDarkShadow
                SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
                UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cShadow
                UserControl.Line (1, He - 3)-(Wi - 2, He - 3), cShadow
                SetPixel .hdc, Wi - 4, He - 4, cShadow
                UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), ShiftColor(cShadow, -&H10)
                UserControl.Line (3, He - 2)-(Wi - 2, He - 2), ShiftColor(cShadow, -&H10)
                SetPixel .hdc, Wi - 2, He - 3, ShiftColor(cShadow, -&H20)
                SetPixel .hdc, Wi - 3, He - 2, ShiftColor(cShadow, -&H20)
                SetPixel .hdc, 2, He - 2, ShiftColor(cShadow, -&H20)
                SetPixel .hdc, 2, He - 3, ShiftColor(cShadow, -&H10)
                SetPixel .hdc, 1, He - 3, ShiftColor(cShadow, -&H10)
                SetPixel .hdc, Wi - 2, 2, ShiftColor(cShadow, -&H20)
                SetPixel .hdc, Wi - 3, 2, ShiftColor(cShadow, -&H10)
                SetPixel .hdc, Wi - 3, 1, ShiftColor(cShadow, -&H10)
            Case 5 'Java
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                .FontBold = True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                UserControl.Line (Wi - 1, 1)-(Wi - 1, He), cHighLight
                UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cHighLight
                DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
                If hasFocus = True Then DrawRectangle (Wi - UserControl.TextWidth(elTex)) \ 2 - 3, (He - UserControl.TextHeight(elTex)) \ 2 - 1, UserControl.TextWidth(elTex) + 6, UserControl.TextHeight(elTex) + 2, &HCC9999, True
                .FontBold = TextFont.Bold
            Case 6 'Netscape
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cShadow, True
                DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                UserControl.Line (Wi - 1, 0)-(Wi - 1, He), ShiftColor(cLight, &H8)
                UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), ShiftColor(cLight, &H8)
                UserControl.Line (0, He - 1)-(Wi, He - 1), ShiftColor(cLight, &H8)
                UserControl.Line (1, He - 2)-(Wi - 1, He - 2), ShiftColor(cLight, &H8)
                If hasFocus = True Then DrawFocusRect .hdc, rc3
             Case 7 'Flat
                Image2.Top = ImgTop + 1
                Image2.Left = ImgLeft + 1
                DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
                DrawRectangle 0, 0, Wi, He, cShadow, True
                UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cHighLight
                UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cHighLight
                If hasFocus = True Then DrawFocusRect .hdc, rc3
          End Select
       End If
    Else
   'DISABLED STATUS
       Select Case MyButtonType
        Case 1 'Windows 16-bit
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            UserControl.Line (1, 0)-(Wi - 1, 0), cDarkShadow
            UserControl.Line (1, He - 1)-(Wi - 1, He - 1), cDarkShadow
            UserControl.Line (0, 1)-(0, He - 1), cDarkShadow
            UserControl.Line (Wi - 1, 1)-(Wi - 1, He - 1), cDarkShadow
            DrawRectangle 1, 1, Wi - 2, He - 2, cHighLight, True
            DrawRectangle 2, 2, Wi - 4, He - 4, cHighLight, True
            UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
            UserControl.Line (Wi - 3, 2)-(Wi - 3, He - 1), cShadow
            UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
            UserControl.Line (2, He - 3)-(Wi - 2, He - 3), cShadow
        Case 2 'Windows 32-bit
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi - 1, He - 1, cHighLight, True
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight, True
            UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cDarkShadow
            UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
            UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cDarkShadow
            UserControl.Line (1, He - 2)-(Wi - 2, He - 2), cShadow
        Case 3 'Windows XP
            stepXP1 = 60 / He
            stepXP2 = 40 / He
            For i = 1 To He
                UserControl.Line (0, i)-(Wi, i), &H9A9A9A - RGB(stepXP1 * i, stepXP1 * i, stepXP2 * i)            '&HFFFFFF
            Next
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            UserControl.Line (2, 0)-(Wi - 2, 0), &H9A9A9A      '&H733C00
            UserControl.Line (2, He - 1)-(Wi - 2, He - 1), &H9A9A9A     ' &H733C00
            UserControl.Line (0, 2)-(0, He - 2), &H9A9A9A     ' &H733C00
            UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), &H9A9A9A     ' &H733C00
            SetPixel .hdc, 1, 1, &H9A9A9A   '&H7B4D10
            SetPixel .hdc, 1, He - 2, &H9A9A9A   '&H7B4D10
            SetPixel .hdc, Wi - 2, 1, &H9A9A9A    '&H7B4D10
            SetPixel .hdc, Wi - 2, He - 2, &H9A9A9A   '&H7B4D10
            DrawRectangle 1, 2, Wi - 2, He - 4, &H464646, True
            UserControl.Line (2, He - 2)-(Wi - 2, He - 2), &H9A9A9A      '&HFE8A71
            UserControl.Line (2, 1)-(Wi - 2, 1), &H9A9A9A     ' &HFFEAD0
            UserControl.Line (1, 2)-(Wi - 1, 2), &H9A9A9A      '&HFAD9BF
        Case 4 'Mac
            DrawRectangle 1, 1, Wi - 2, He - 2, cLight
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            UserControl.Line (2, 0)-(Wi - 2, 0), cDarkShadow
            UserControl.Line (2, He - 1)-(Wi - 2, He - 1), cDarkShadow
            UserControl.Line (0, 2)-(0, He - 2), cDarkShadow
            UserControl.Line (Wi - 1, 2)-(Wi - 1, He - 2), cDarkShadow
            SetPixel .hdc, 1, 1, cDarkShadow
            SetPixel .hdc, 1, He - 2, cDarkShadow
            SetPixel .hdc, Wi - 2, 1, cDarkShadow
            SetPixel .hdc, Wi - 2, He - 2, cDarkShadow
            SetPixel .hdc, 1, 2, cFace
            SetPixel .hdc, 2, 1, cFace
            UserControl.Line (3, 2)-(Wi - 3, 2), cHighLight
            UserControl.Line (2, 2)-(2, He - 3), cHighLight
            SetPixel .hdc, 3, 3, cHighLight
            UserControl.Line (Wi - 3, 1)-(Wi - 3, He - 3), cFace
            UserControl.Line (1, He - 3)-(Wi - 3, He - 3), cFace
            SetPixel .hdc, Wi - 4, He - 4, cFace
            UserControl.Line (Wi - 2, 3)-(Wi - 2, He - 2), cShadow
            UserControl.Line (3, He - 2)-(Wi - 2, He - 2), cShadow
            SetPixel .hdc, Wi - 3, He - 3, cShadow
            SetPixel .hdc, 2, He - 2, cFace
            SetPixel .hdc, 2, He - 3, cLight
            SetPixel .hdc, Wi - 2, 2, cFace
            SetPixel .hdc, Wi - 3, 2, cLight
        Case 5 'Java
            .FontBold = True
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi, He, cShadow, True
            .FontBold = TextFont.Bold
        Case 6 'Netscape
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi, He, ShiftColor(cLight, &H8), True
            DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cLight, &H8), True
            UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
            UserControl.Line (Wi - 2, 1)-(Wi - 2, He - 1), cShadow
            UserControl.Line (0, He - 1)-(Wi, He - 1), cShadow
            UserControl.Line (1, He - 2)-(Wi - 1, He - 2), cShadow
        Case 7 'Flat
            SetTextColor .hdc, cHighLight
            DrawText .hdc, elTex, Len(elTex), rc2, DT_CENTERABS
            SetTextColor .hdc, cShadow
            DrawText .hdc, elTex, Len(elTex), rc, DT_CENTERABS
            DrawRectangle 0, 0, Wi, He, cHighLight, True
            UserControl.Line (Wi - 1, 0)-(Wi - 1, He), cShadow
            UserControl.Line (0, He - 1)-(Wi - 1, He - 1), cShadow
      End Select
    End If
  End With
  hasFocus = preFocusValue
End Sub

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)
' Disegno dei rettangoli utilizzando le linee
  Dim bRect As RECT
  Dim hBrush As Long
  Dim Ret As Long
  bRect.Left = x
  bRect.Top = y
  bRect.Right = x + Width
  bRect.Bottom = y + Height
  hBrush = CreateSolidBrush(Color)
  If OnlyBorder = False Then
     Ret = FillRect(UserControl.hdc, bRect, hBrush)
  Else
     Ret = FrameRect(UserControl.hdc, bRect, hBrush)
  End If
  Ret = DeleteObject(hBrush)
End Sub

Private Sub SetColors()
' Setta i colori in base allo stile
  If MyColorType = Custom Then
     cFace = BackC
     cText = ForeC
     cShadow = ShiftColor(cFace, -&H40)
     cLight = ShiftColor(cFace, &H1F)
     cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
     cDarkShadow = ShiftColor(cFace, -&HC0)
  ElseIf MyColorType = [Force Standard] Then
     cFace = &HC0C0C0
     cShadow = &H808080
     cLight = &HDFDFDF
     cDarkShadow = &H0
     cHighLight = &HFFFFFF
     cText = &H0
  Else
    'se MyColorType = 1 o non e' settato utilizza i colori windows
     cFace = GetSysColor(COLOR_BTNFACE)
     cShadow = GetSysColor(COLOR_BTNSHADOW)
     cLight = GetSysColor(COLOR_BTNLIGHT)
     cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
     cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
     cText = GetSysColor(COLOR_BTNTEXT)
  End If
End Sub

Private Sub MakeRegion()
' Creazione delle regions dove 'tagliare' il controllo
' utilizzando le trasparenze
  Dim rgn1 As Long, rgn2 As Long
  DeleteObject rgnNorm
  rgnNorm = CreateRectRgn(0, 0, Wi, He)
  rgn2 = CreateRectRgn(0, 0, 0, 0)
  Select Case MyButtonType
    Case 1 'Windows 16-bit
        rgn1 = CreateRectRgn(0, 0, 1, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
    Case 3, 4 'Windows XP and Mac
        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
    Case 5 'Java
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
  End Select
  DeleteObject rgn2
End Sub

Private Sub SetAccessKeys()
  Dim ampersandPos As Long
  If Len(elTex) > 1 Then
     ampersandPos = InStr(1, elTex, "&", vbTextCompare)
     If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
        If Mid(elTex, ampersandPos + 1, 1) <> "&" Then
            UserControl.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
        Else
            ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
            If Mid(elTex, ampersandPos + 1, 1) <> "&" Then
                UserControl.AccessKeys = LCase(Mid(elTex, ampersandPos + 1, 1))
            Else
                UserControl.AccessKeys = ""
            End If
        End If
     Else
        UserControl.AccessKeys = ""
     End If
  Else
     UserControl.AccessKeys = ""
  End If
End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long) As Long
' Inserimento e colori
  Dim Red As Long, Blue As Long, Green As Long
  Blue = ((Color \ &H10000) Mod &H100) + Value
  Green = ((Color \ &H100) Mod &H100) + Value
  Red = (Color And &HFF) + Value
' check red
  If Red < 0 Then
     Red = 0
  ElseIf Red > 255 Then
     Red = 255
  End If
' check green
  If Green < 0 Then
     Green = 0
  ElseIf Green > 255 Then
     Green = 255
  End If
' check blue
  If Blue < 0 Then
     Blue = 0
  ElseIf Blue > 255 Then
     Blue = 255
  End If
  ShiftColor = RGB(Red, Green, Blue)
End Function
