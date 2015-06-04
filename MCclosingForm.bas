Attribute VB_Name = "McClosingForm"
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const WS_EX_TOPMOST = &H8&
Private Const WS_BORDER = &H800000
Private Const WS_SYSMENU = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

Private Const SW_SHOWNORMAL = 1
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOCOPYBITS = &H100

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long


Public Sub MCCloseForm(FormToClose As Form, process As Variant)
  Dim Hwndnew As Long
  Dim R As RECT
  Dim Rwidth As Integer
  Dim Rheight As Integer
  
  GetWindowRect FormToClose.hwnd, R
  Rwidth = R.Right - R.Left
  Rheight = R.Bottom - R.Top
  
    
    Hwndnew = CreateWindowEx(0, "static", "", WS_POPUPWINDOW Or WS_EX_TOPMOST, R.Top, R.Left, Rwidth, Rheight, 0, 0, App.hInstance, ByVal 0&)
    ShowWindow Hwndnew, SW_SHOWNORMAL
    Unload FormToClose
    R.Left = R.Left + 1
    SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
  
    'SetCursorPos 0, Screen.Height / Screen.TwipsPerPixelY / 2
    'mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
 
    If process = "Rnd" Then
    Randomize
    process = Int(Rnd * 16)
    End If
    
    
'    Select Case process
'    Case 1
'            Do
'            If Rheight < 3 Then Exit Do
'            R.Left = R.Left + 1
'            R.Top = R.Top + 1
'            Rwidth = Rwidth - 2
'            Rheight = Rheight - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'
 '   Case 2
            Do
            If Rheight < 3 Then Exit Do
            R.Left = R.Left + 1
            R.Top = R.Top + 1
            Rwidth = Rwidth - 2
            Rheight = Rheight - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents
            Loop
            
            Do
            If Rwidth < 0 Then Exit Do
            R.Left = R.Left + 1
            Rwidth = Rwidth - 2
            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
            DoEvents
            Loop
'    Case 3
'            Do
'            If Rwidth < 3 Then Exit Do
'            R.Left = R.Left + 1
'            Rwidth = Rwidth - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'
'            Do
'            If Rheight < 0 Then Exit Do
'            R.Top = R.Top + 1
'            Rheight = Rheight - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'    Case 4
'            Do
'            If Rheight < 3 Then Exit Do
'            R.Top = R.Top + 1
'            Rheight = Rheight - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'
'            Do
'            If Rwidth < 0 Then Exit Do
'            R.Left = R.Left + 1
'            Rwidth = Rwidth - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'     Case 5
'            Do
'            If Rwidth < 0 Then Exit Do
'            R.Left = R.Left + 1
'            Rwidth = Rwidth - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'    Case 6
'            Do
'            If Rwidth < 0 Then Exit Do
'            Rwidth = Rwidth - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'    Case 7
'            Do
'            If Rheight < 0 Then Exit Do
'            Rheight = Rheight - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'    Case 8
'            Do
'            If Rheight < 0 Then Exit Do
'            R.Top = R.Top + 1
'            Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW Or SWP_NOCOPYBITS
'            DoEvents
'            Loop
'    Case 9
'            Do
'            If R.Left > Screen.Width / Screen.TwipsPerPixelX Then Exit Do
'            R.Left = R.Left + 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'    Case 10
'            Do
'            If Abs(R.Left) > Rwidth Then Exit Do
'            R.Left = R.Left - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'     Case 11
'            Do
'            If Abs(R.Top) > Rheight Then Exit Do
'            'R.Left = R.Left - 1
'            R.Top = R.Top - 1
'            'Rwidth = Rwidth - 1
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'    Case 12
'            Do
'            If R.Top > Screen.Height / Screen.TwipsPerPixelY Then Exit Do
'            'R.Left = R.Left - 1
'            R.Top = R.Top + 1
'            'Rwidth = Rwidth - 1
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents 'enable redrawing
'            Loop
'    Case 13
'            Do
'            If Abs(R.Top) > Rheight Then Exit Do
'            R.Left = R.Left + 1
'            R.Top = R.Top - 1
'            'Rwidth = Rwidth - 1
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents 'enable redrawing
'            Loop
'    Case 14
'            Do
'            If Abs(R.Top) > Rheight Then Exit Do
'            R.Left = R.Left - 1
'            R.Top = R.Top - 1
'            'Rwidth = Rwidth - 1
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'    Case 15
'            Do
'            If R.Top > Screen.Height / Screen.TwipsPerPixelY Then Exit Do
'            R.Left = R.Left - 1
'            R.Top = R.Top + 1
'            'Rwidth = Rwidth - 1
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'    Case 16
'            Do
'            If R.Top > Screen.Height / Screen.TwipsPerPixelY Then Exit Do
'            R.Left = R.Left + 1
'            R.Top = R.Top + 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'    Case "Min"
'            Do
'            If Rheight < 25 Then Exit Do
'            Rheight = Rheight - 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'
'            Do
'            If R.Top > (Screen.Height / Screen.TwipsPerPixelY) - 40 Then Exit Do
'            R.Top = R.Top + 2
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'
'            Do
'            If Rwidth < 26 Then Exit Do
'            Rwidth = Rwidth - 4
'            R.Left = R.Left + 2
'            R.Top = R.Top - 2
'
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'
'            Do
'            If R.Top > (Screen.Height / Screen.TwipsPerPixelY) - 40 Then Exit Do
'            'Rwidth = Rwidth - 2
'            R.Left = R.Left + 2
'            R.Top = R.Top + 2
'
'            'Rheight = Rheight - 1
'            SetWindowPos Hwndnew, HWND_TOPMOST, R.Left, R.Top, Rwidth, Rheight, SWP_SHOWWINDOW
'            DoEvents
'            Loop
'
'
'
'    Case Else
'    End Select
'
    DestroyWindow Hwndnew
End Sub



