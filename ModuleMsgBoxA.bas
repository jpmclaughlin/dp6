Attribute VB_Name = "ModuleMsgBoxA"
Option Explicit

Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Const MB_APPLMODAL = &H0&
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_NOFOCUS = &H8000&

Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_DEFBUTTON2 = &H100&
Public Const MB_DEFBUTTON3 = &H200&

Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&

Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONERROR = MB_ICONHAND
Public Const MB_ICONSTOP = MB_ICONHAND

'Return Value

Public Const IDABORT = 3
Public Const IDCANCEL = 2
Public Const IDIGNORE = 5
Public Const IDNO = 7
Public Const IDOK = 1
Public Const IDYES = 6
Public Const IDRETRY = 4

'Se hWnd = 0 la finestra di dialogo fa riferimento al desktop. L'applicazione non si blocca

Public Function MsgBoxA( _
    ByVal hWnd As Long, _
    ByVal Prompt As String, _
    Optional ByVal Title As String = "", _
    Optional ByVal Flag_MB_ICON As Long = 0, _
    Optional ByVal Flag_MB_BUTTON As Long = MB_OK, _
    Optional ByVal Flag_MB_DEFAULTBUTTON As Long = MB_DEFBUTTON1, _
    Optional ByVal Other_Flags As Long = 0) As Long

    Dim Flags As Long
    
    Flags = Flag_MB_ICON Or _
            Flag_MB_BUTTON Or _
            Flag_MB_DEFAULTBUTTON Or _
            Other_Flags

    MsgBoxA = MessageBox(hWnd, Prompt, Title, Flags)

End Function

