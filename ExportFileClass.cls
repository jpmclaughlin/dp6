VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportFileClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function EXPORT_BundleLog(ByVal inString As String) As Boolean
    Dim f As Integer
    
    f = FreeFile
    Err.Clear
    On Error Resume Next
    EXPORT_BundleLog = False
    If Param.GetNumber("Par228_Exports") < 1 Then
'       If Dir("..\target\BundleLOG_prev.Dat") <> "" Then
'          Kill "..\target\BundleLOG_prev.Dat"
'       End If
'       Name "..\target\BundleLOG.Dat" As "..\target\BundleLOG_prev.Dat"
       Open "..\target\BundleLOG.Dat" For Output As #f
'       Param.SetNumber "Par228_Exports", 1
    Else
       Open "..\target\BundleLOG.Dat" For Append As #f
'       Param.SetNumber "Par228_Exports", Param.GetNumber("Par228_Exports") + 1
    End If
    Print #f, inString
    Close #f
    EXPORT_BundleLog = Err <> 0
End Function

Function Increase_Exports_number() As Long
    If Param.GetNumber("Par228_Exports") > 15000 Then
       If Dir("..\target\BundleLOG_prev.Dat") <> "" Then
          Kill "..\target\BundleLOG_prev.Dat"
       End If
       Name "..\target\BundleLOG.Dat" As "..\target\BundleLOG_prev.Dat"
       
       If Dir("..\target\Bund_Data_prev.Dat") <> "" Then
          Kill "..\target\Bund_Data_prev.Dat"
       End If
       Name "..\target\Bund_Data.Dat" As "..\target\Bund_Data_prev.Dat"
       
       Param.SetNumber "Par228_Exports", 0
    Else
        Param.SetNumber "Par228_Exports", Param.GetNumber("Par228_Exports") + 1
    End If
    
    Increase_Exports_number = Param.GetNumber("Par228_Exports")
End Function

Function SaveBundleData(ByVal inString As String) As Boolean
    Dim f As Integer
    
    f = FreeFile
    Err.Clear
    On Error Resume Next
    SaveBundleData = False
    If Param.GetNumber("Par228_Exports") < 1 Then
'       If Dir("..\target\Bund_Data_prev.Dat") <> "" Then
'          Kill "..\target\Bund_Data_prev.Dat"
'       End If
'       Name "..\target\Bund_Data.Dat" As "..\target\Bund_Data_prev.Dat"
       Open "..\target\Bund_Data.Dat" For Output As #f
'       Param.SetNumber "Par228_Exports", 0
    Else
       Open "..\target\Bund_Data.Dat" For Append As #f
'       Param.SetNumber "Par228_Exports", Param.GetNumber("Par228_Exports") + 1
    End If
    Print #f, inString
    Close #f
    SaveBundleData = Err <> 0
End Function

Function UniqueKeyAdd(ByVal inMac As String) As String
    Dim f As Integer
    Dim inString As String
    Dim number As Long
    
    f = FreeFile
    Err.Clear
    On Error Resume Next
    Open "..\target\Sland_Ser_Num.Dat" For Input As #f
    Input #f, inString
    Close #f
    number = Val(Right(inString, 6))
    number = number + 1
    If number > 999999 Then number = 1
    UniqueKeyAdd = inMac & Format(number, "000000")
    inString = UniqueKeyAdd
    Open "..\target\Sland_Ser_Num.Dat" For Output As #f
    Print #f, inString
    Close #f
End Function

Function UniqueKeyread() As String
    Dim f As Integer
    Dim inString As String
    Dim number As Long
    
    f = FreeFile
    Err.Clear
    On Error Resume Next
    Open "..\target\Sland_Ser_Num.Dat" For Input As #f
    Input #f, inString
    UniqueKeyread = inString
    Close #f
End Function
