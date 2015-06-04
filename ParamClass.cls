VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParamClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' costanti generali per reperire i parametri
Private Const ParamDataBaseName As String = "..\target\parameters.mdb"
Private Const BackupDataBaseName As String = "..\databackup\parameters.mdb"
Private Const BitTableName As String = "OptionsQuery"
Private Const NumbersTableName As String = "NumbersQuery"
Private Const TextTableName As String = "LanguageQuery"
' nome campo testo (dipende dalla lingua)
Private TextFieldName As String

Private dbs As Database
Private rstString As Recordset
Private strQuery As String

Private Sub Class_Initialize()
   
    ' verifica esistenza e integrità del file parametri
    On Error Resume Next
        Set dbs = OpenDatabase(ParamDataBaseName)
        dbs.Close
        If Err.Number <> 0 Then
            ' file danneggiato: ripristina usando la copia di sicurezza
            If Dir(ParamDataBaseName) <> "" Then Kill ParamDataBaseName
            FileCopy BackupDataBaseName, ParamDataBaseName
        Else
            ' file non danneggiato: aggiorna la copia di sicurezza
            If Dir(BackupDataBaseName) <> "" Then Kill BackupDataBaseName
            FileCopy ParamDataBaseName, BackupDataBaseName
        End If
    On Error GoTo 0
    
    
    ' apertura database
    Set dbs = OpenDatabase(ParamDataBaseName)
    TextFieldName = "ItalianString"
    If Bit("ItalianText") Then TextFieldName = "ItalianString"
    If Bit("EnglishText") Then TextFieldName = "EnglishString"
    If Bit("SpanishText") Then TextFieldName = "SpanishString"
    If Bit("FrenchText") Then TextFieldName = "FrenchString"
    If Bit("GermanText") Then TextFieldName = "GermanString"
        
End Sub

Property Get DataBaseName() As String
    DataBaseName = ParamDataBaseName
End Property

Property Get LanguageTableName() As String
    LanguageTableName = TextTableName
End Property

Property Get Text(Identifier As String) As String
    '                       campo                     tabella
    strQuery = "SELECT " & TextFieldName & " FROM " & TextTableName & " WHERE ID='" _
                & Identifier & "';"
    Set rstString = dbs.OpenRecordset(strQuery, dbOpenDynaset)
    If Not rstString.EOF Then
        If rstString.Fields(TextFieldName) <> "" Then
            Text = rstString.Fields(TextFieldName)
        Else
            Text = "*"
        End If
    Else
        Text = "***"
        MsgBox "Parameter """ & Identifier & """ not present in table """ & TextTableName & """", , "Parameter database"
    End If
End Property


Property Get Number(Identifier As String) As Double
    '                  campo      tabella
    strQuery = "SELECT Value FROM " & NumbersTableName & " WHERE ID='" _
     & Identifier & "';"
    Set rstString = dbs.OpenRecordset(strQuery, dbOpenDynaset)
    If Not rstString.EOF Then
        Number = rstString.Fields("Value")
    Else
        Number = 1#
        MsgBox "Parameter """ & Identifier & """ not present in table """ & NumbersTableName & """", , "Parameter database"
    End If
End Property

Property Let Number(Identifier As String, value As Double)
    '                  campo          tabella
    strQuery = "SELECT Value FROM " & NumbersTableName & " WHERE ID='" _
     & Identifier & "';"
    Set rstString = dbs.OpenRecordset(strQuery, dbOpenDynaset)
    If Not rstString.EOF Then
        rstString.Edit
        rstString.Fields("Value") = value
        rstString.Update
    Else
        MsgBox "Parameter """ & Identifier & """ not present in table """ & NumbersTableName & """", , "Parameter database"
    End If
End Property

Property Get Bit(Identifier As String) As Boolean
    '                  campo          tabella
    strQuery = "SELECT Bit   FROM " & BitTableName & " WHERE ID='" _
     & Identifier & "';"
    Set rstString = dbs.OpenRecordset(strQuery, dbOpenDynaset)
    If Not rstString.EOF Then
        Bit = rstString.Fields("Bit")
    Else
        Bit = False
        MsgBox "Parameter """ & Identifier & """ not present in table """ & BitTableName & """", , "Parameter database"
    End If
End Property


Property Let Bit(Identifier As String, value As Boolean)
    '                  campo          tabella
    strQuery = "SELECT Bit   FROM " & BitTableName & " WHERE ID='" _
     & Identifier & "';"
    Set rstString = dbs.OpenRecordset(strQuery, dbOpenDynaset)
    If Not rstString.EOF Then
        rstString.Edit
        rstString.Fields("Bit") = value
        rstString.Update
    Else
        MsgBox "Parameter """ & Identifier & """ not present in table """ & BitTableName & """", , "Parameter database"
    End If
End Property


