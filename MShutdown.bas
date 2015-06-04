Attribute VB_Name = "MShutdown"
Option Explicit
Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Const EWX_SHUTDOWN As Long = 1
Const EWX_FORCE As Long = 4
Const EWX_REBOOT As Long = 2
Const EWX_POWEROFF As Long = 8

Declare Function ExitWindowsEx Lib "User32.dll" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

Declare Function GetCurrentProcess Lib "Kernel32.dll" () As Long

'Dichiarazioni per acquisire i diritti necessari all'arresto/riavvio della macchina.
Declare Function OpenProcessToken Lib "Advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Declare Function LookupPrivilegeValue Lib "Advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Declare Function AdjustTokenPrivileges Lib "Advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Sub AdjustToken()
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long, hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES, tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long

    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle

    'Get  LUID per il privilegio di chiusura.
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid

    With tkp
        tkp.PrivilegeCount = 1    ' un privilegio settato
        tkp.TheLuid = tmpLuid
        tkp.Attributes = SE_PRIVILEGE_ENABLED
    End With

    'abilita il token per il processo corrente
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded

End Sub

